# -*- coding: utf-8 -*-
"""Purpose: validate Space payloads for Loadwise's full-parameter coordination API.
Behaviors: every named parameter, exact IDs, bounded values and chunked pushes.
Revit API: none; the host collects JSON-compatible values in its API callback.
Design: validate before HTTP. SpacesSnapshot is the main coordination adapter;
SpacesPayload retains the separate number/name/geometry load-room contract.
"""
import math
import uuid

try:
    STRING_TYPES = (basestring,)
except NameError:
    STRING_TYPES = (str,)

NUMBER_TYPES = (int, float)
try:
    NUMBER_TYPES = NUMBER_TYPES + (long,)
except NameError:
    pass


def _text(value, label, maximum):
    if value is None:
        return ''
    if not isinstance(value, STRING_TYPES):
        raise ValueError('{0} must be text.'.format(label))
    value = value.strip()
    if len(value) > maximum:
        raise ValueError('{0} is longer than {1} characters.'.format(label, maximum))
    return value


def _measure(value, label, maximum):
    if value is None:
        return None
    if (isinstance(value, bool) or not isinstance(value, NUMBER_TYPES)
            or math.isnan(value) or math.isinf(value) or value < 0 or value > maximum):
        raise ValueError('{0} must be a finite number between 0 and {1}.'.format(label, maximum))
    return round(float(value), 4)


def SpacesPayload(rows, space_type_id=None):
    """Return the request body for sync_revit_spaces (without projectId)."""
    spaces = []
    seen = set()
    for row in rows:
        unique_id = _text(row.get('uniqueId'), 'Space UniqueId', 200)
        if not unique_id:
            raise ValueError('Every Space needs its Revit UniqueId.')
        if unique_id in seen:
            raise ValueError('Space {0} appears twice.'.format(unique_id))
        seen.add(unique_id)
        element_id = row.get('id')
        if element_id is not None and not isinstance(element_id, STRING_TYPES):
            raise ValueError('Element IDs must be strings so large IDs keep their exact value.')
        space = {'uniqueId': unique_id,
                 'number': _text(row.get('number'), 'Number', 50),
                 'name': _text(row.get('name'), 'Name', 200),
                 'level': _text(row.get('level'), 'Level', 100)}
        if element_id:
            space['elementId'] = _text(element_id, 'Element ID', 30)
        area = _measure(row.get('areaValue'), 'Area', 1000000)
        height = _measure(row.get('heightValue'), 'Height', 200)
        if area is not None:
            space['area'] = area
        if height is not None:
            space['ceilingHeight'] = height
        spaces.append(space)
    request = {'spaces': spaces}
    if space_type_id:
        request['spaceTypeId'] = space_type_id
    return request


def SyncSpaces(client, project_id, rows, space_type_id=None):
    """Push Spaces to a Loadwise project. Missing Spaces are reported, never deleted."""
    request = SpacesPayload(rows, space_type_id)
    request['projectId'] = project_id
    return client.sync_revit_spaces(request)


# ---------------------------------------------------------------- full-parameter snapshot (push_revit_spaces)
# Rows: {'uniqueId', 'id' (element ID string), 'number', 'name', 'level',
#        'parameters': {parameter name: {'v': value in display units, 's': display string}}}.
# Values are numbers already converted to the model's display units (ft², ft³, CFM, Btu/h ...), strings,
# booleans or None. The catalog lists each parameter once: {'name', 'storageType', 'unit', 'group',
# 'isShared', 'isType'}.

def _value(value, label):
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, NUMBER_TYPES):
        if value != value or value in (float('inf'), float('-inf')):
            raise ValueError('{0} must be a finite number.'.format(label))
        return round(float(value), 6) if isinstance(value, float) else value
    if isinstance(value, STRING_TYPES):
        if len(value) > 2000:
            raise ValueError('{0} is longer than 2000 characters.'.format(label))
        return value
    raise ValueError('{0} must be a number, text, a Yes/No value or empty.'.format(label))


def _parameters(values, label):
    if values is None:
        return {}
    if not isinstance(values, dict):
        raise ValueError('{0} parameters must be a dictionary.'.format(label))
    if len(values) > 2000:
        raise ValueError('{0} has more than 2000 parameters.'.format(label))
    result = {}
    for name, item in values.items():
        if not isinstance(name, STRING_TYPES) or not name.strip() or len(name) > 200:
            raise ValueError('{0} has a parameter without a usable name.'.format(label))
        where = '{0} parameter {1}'.format(label, name)
        if not isinstance(item, dict) or set(item) - set(('v', 's')):
            raise ValueError('{0} must be {{"v": value, "s": display text}}.'.format(where))
        entry = {'v': _value(item.get('v'), where)}
        if item.get('s') is not None:
            entry['s'] = _text(item.get('s'), where + ' display', 2000)
        result[name] = entry
    return result


def _catalog(items):
    catalog = []
    seen = set()
    for item in items or ():
        if not isinstance(item, dict):
            raise ValueError('Every catalog entry must be a dictionary.')
        name = _text(item.get('name'), 'Parameter name', 200)
        if not name or name in seen:
            raise ValueError('Catalog parameter names must be present and unique.')
        seen.add(name)
        entry = {'name': name, 'storageType': _text(item.get('storageType'), 'Storage type', 20),
                 'unit': _text(item.get('unit'), 'Unit', 50), 'group': _text(item.get('group'), 'Group', 100),
                 'isShared': bool(item.get('isShared')), 'isType': bool(item.get('isType'))}
        catalog.append(entry)
    return catalog


def SpacesSnapshot(rows):
    """Validate every row and return the push_revit_spaces 'spaces' list. Raises ValueError before any request."""
    spaces = []
    seen = set()
    for row in rows:
        unique_id = _text(row.get('uniqueId'), 'Space UniqueId', 200)
        if not unique_id:
            raise ValueError('Every Space needs its Revit UniqueId.')
        if unique_id in seen:
            raise ValueError('Space {0} appears twice.'.format(unique_id))
        seen.add(unique_id)
        element_id = row.get('id')
        if element_id is not None and not isinstance(element_id, STRING_TYPES):
            raise ValueError('Element IDs must be strings so large IDs keep their exact value.')
        space = {'uniqueId': unique_id,
                 'number': _text(row.get('number'), 'Number', 50),
                 'name': _text(row.get('name'), 'Name', 200),
                 'level': _text(row.get('level'), 'Level', 100),
                 'parameters': _parameters(row.get('parameters'), 'Space ' + unique_id)}
        if element_id:
            space['elementId'] = _text(element_id, 'Element ID', 30)
        spaces.append(space)
    return spaces


def PushSpaces(client, project_id, rows, catalog=None, chunk_size=200, sync_id=None):
    """Send every Space with all its parameters, in chunks that share one sync ID.

    Only the last chunk is marked final; Loadwise then flags Spaces missing from this push (never deletes them).
    Returns the summed created/updated/unchanged counts, the newly missing UniqueIds and the sync status."""
    spaces = SpacesSnapshot(rows)
    if not spaces:
        raise ValueError('No Spaces to push. An empty push would mark every Space as missing.')
    if chunk_size < 1 or chunk_size > 1000:
        raise ValueError('chunk_size must be between 1 and 1000.')
    parameters = _catalog(catalog)
    sync_id = sync_id or str(uuid.uuid4())
    totals = {'created': 0, 'updated': 0, 'unchanged': 0, 'newlyMissing': [], 'sync': None, 'chunks': 0}
    starts = list(range(0, len(spaces), chunk_size))
    for index, start in enumerate(starts):
        request = {'projectId': project_id, 'syncId': sync_id, 'spaces': spaces[start:start + chunk_size],
                   'final': index == len(starts) - 1}
        if index == 0 and parameters:
            request['parameters'] = parameters
        result = client.push_revit_spaces(request)
        for key in ('created', 'updated', 'unchanged'):
            totals[key] += result.get(key, 0)
        totals['newlyMissing'] = result.get('newlyMissing', [])
        totals['sync'] = result.get('sync')
        totals['chunks'] += 1
    return totals
