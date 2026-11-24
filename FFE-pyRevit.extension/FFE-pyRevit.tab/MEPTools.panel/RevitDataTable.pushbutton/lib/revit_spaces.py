# -*- coding: utf-8 -*-
"""Convert collected Revit Space rows into a Loadwise sync_revit_spaces request.

No Revit API calls: the host collects plain dictionaries on the Revit thread, then
calls SyncSpaces off-thread. Revit owns number, name, level, area and height;
Loadwise load inputs are never sent from Revit.

Each row needs 'uniqueId' (Space.UniqueId) and may carry 'id' (element ID as a
string), 'number', 'name', 'level', 'areaValue' (ft², Revit internal units) and
'heightValue' (ft, Space.UnboundedHeight in internal units).
"""
import math

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
