# -*- coding: utf-8 -*-
"""Convert the existing JSON snapshot into an explicit repository collection.

No Revit API access. The host must collect capabilities on the Revit thread:
capabilities[space.UniqueId][parameter_identity] = {
    'writable': not parameter.IsReadOnly, 'canUnset': <supported clear operation>,
    'hasValue': parameter.HasValue, 'referenceUniqueId': <resolved ElementId> }
The default coverage is partial; callers explicitly attest to a complete scan.
"""
import json
import math

BASES = (
    ('number', 'Number', 'String', None), ('name', 'Name', 'String', None),
    ('areaValue', 'Area', 'Double', 'autodesk.spec.aec:area-2.0.0'),
    ('heightValue', 'Unbounded Height', 'Double', 'autodesk.spec.aec:length-2.0.0'),
    ('level', 'Level', 'ElementId', None), ('phase', 'Phase', 'ElementId', None),
)


def from_snapshot(snapshot, capabilities=None, coverage=None, references=None):
    capabilities = capabilities or {}
    references = references or {}
    parameters = []
    sources = []
    for field, name, storage, spec in BASES:
        definition = {'identity': 'base:' + field, 'name': name, 'storageType': storage}
        if spec:
            definition['specTypeId'] = spec
        parameters.append(definition)
        sources.append((definition, field, True))
    for original in snapshot.get('parameters', []):
        if original.get('baseField'):
            continue
        # Only values actually collected need a definition. Never substitute names.
        if not any(original.get('key') in row.get('parameterValues', {}) for row in snapshot.get('rows', [])):
            continue
        if not original.get('identity'):
            raise ValueError('A collected parameter is missing its stable identity: ' + original.get('name', ''))
        storage = 'Boolean' if original.get('isBoolean') else original.get('storageType')
        definition = {'identity': original['identity'], 'name': original['name'], 'storageType': storage}
        if storage == 'Double':
            spec = original.get('specTypeId') or original.get('dataTypeId')
            if not spec:
                raise ValueError('A numeric parameter is missing specTypeId: ' + original['identity'])
            definition['specTypeId'] = spec
        if original.get('unitTypeId'):
            definition['unitTypeId'] = original['unitTypeId']
        parameters.append(definition)
        sources.append((definition, original['key'], False))
    if len(set(d['identity'] for d in parameters)) != len(parameters):
        raise ValueError('Duplicate parameter identities in collection.')
    spaces = []
    seen = set()
    warnings = list(snapshot.get('issues', []))
    for row in snapshot.get('rows', []):
        unique_id = row.get('uniqueId')
        if not unique_id or unique_id in seen:
            raise ValueError('Every Space needs a distinct Revit UniqueId.')
        if not isinstance(row.get('id'), (type(''), type(u''))):
            raise ValueError('Revit element IDs must remain strings.')
        seen.add(unique_id)
        warnings.extend({'spaceId': unique_id, 'message': warning} for warning in row.get('warnings', []))
        cells = []
        for definition, field, is_base in sources:
            identity = definition['identity']
            capability = capabilities.get(unique_id, {}).get(identity, {})
            values = row if is_base else row.get('parameterValues', {})
            if not is_base and field not in values:
                continue
            raw = values.get(field)
            if is_base and field in ('level', 'phase'):
                item_id = row.get(field + 'Id')
                catalog = snapshot.get('levels' if field == 'level' else 'phases', [])
                matches = [x for x in catalog if x.get('id') == item_id and x.get('uniqueId')]
                raw = {'uniqueId': matches[0]['uniqueId']} if len(matches) == 1 else None
            storage = definition['storageType']
            if storage == 'ElementId' and raw is not None and not isinstance(raw, dict):
                reference_id = capability.get('referenceUniqueId') or references.get(str(raw))
                if reference_id:
                    raw = {'uniqueId': reference_id}
                elif str(raw).startswith('-') and str(raw)[1:].isdigit():
                    raw = {'sentinel': str(raw)}
                else:
                    raw = None
            if storage == 'Boolean' and raw is not None:
                if raw not in (True, False, 0, 1):
                    raise ValueError('Invalid Yes/No value.')
                raw = bool(raw)
            if isinstance(raw, float) and (math.isnan(raw) or math.isinf(raw)):
                raw = None
            entry = {'parameter': identity}
            if raw is None and capability.get('hasValue') is not False:
                entry['availability'] = 'unavailable'
            else:
                entry.update({'availability': 'collected', 'state': {'state': 'unset'} if raw is None else {'state': 'value', 'value': raw},
                              'writable': capability.get('writable') is True and identity not in ('base:areaValue', 'base:level', 'base:phase'),
                              'canUnset': capability.get('canUnset') is True})
                text = row.get('parameterTexts', {}).get(field) if not is_base else row.get({'areaValue': 'areaText', 'heightValue': 'heightText'}.get(field, field))
                if isinstance(text, (type(''), type(u''))):
                    entry['displayText'] = text
            cells.append(entry)
        spaces.append({'uniqueId': unique_id, 'elementId': row['id'], 'context': {}, 'cells': cells})
    selected_coverage = dict(coverage or {'spaces': 'partial', 'successful': True})
    skipped = snapshot.get('skippedCount', 0)
    if skipped:
        selected_coverage['successful'] = False
    result = {'parameters': parameters, 'spaces': spaces, 'coverage': selected_coverage,
              'skippedCount': skipped, 'warnings': warnings,
              'sourceCollectedAt': snapshot.get('refreshedAt')}
    # Deep copy and reject non-JSON/non-finite values before handing data to HTTP.
    return json.loads(json.dumps(result, allow_nan=False))
