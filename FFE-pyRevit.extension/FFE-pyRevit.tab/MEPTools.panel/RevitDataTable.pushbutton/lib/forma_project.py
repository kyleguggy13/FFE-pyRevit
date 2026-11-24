# -*- coding: utf-8 -*-
"""Autodesk Forma project discovery and startup connection, with injectable APIs.

ReadFormaContext is called only in a valid Revit API context. Everything returned
is plain JSON; CheckFormaProject/CreateFormaProject can then run off-thread.
No credentials are obtained from Revit, and no Space data is written here.
"""
import re
import uuid
try:
    from urllib.parse import unquote
except ImportError:
    from urllib import unquote


def NormalizeProjectId(value):
    text = str(value or '').strip()
    if text.lower().startswith('b.'):
        text = text[2:]
    parsed = uuid.UUID(text)
    if parsed.int == 0:
        raise ValueError('Autodesk Forma project ID is empty.')
    return str(parsed)


def ProjectNameFromPath(path):
    """Read only explicitly named Autodesk cloud paths, never Project Information."""
    match = re.match(r'^(?:BIM\s*360|Autodesk\s+Docs|Autodesk\s+Forma|Forma)://([^/\\]+)[/\\].+',
                     path or '', re.IGNORECASE)
    if not match:
        return None
    name = unquote(match.group(1).encode('utf-8')).strip()
    if not isinstance(name, type(u'')):
        name = name.decode('utf-8')
    try:
        NormalizeProjectId(name)
        return None  # GUID-only paths do not contain a human-readable project name.
    except (ValueError, TypeError, AttributeError):
        pass
    return name if name and len(name) <= 200 else None


def ReadFormaContext(document, db=None):
    """Capture the real Forge Data Management ID while Revit grants API access."""
    try:
        if not document.IsModelInCloud:
            return None
    except AttributeError:
        return None
    try:
        project_id = NormalizeProjectId(document.GetProjectId())
    except Exception:
        return {'available': False, 'reason': 'Revit could not read the Autodesk Forma project ID. '
                'Check the cloud connection, then close and reopen HVAC Data.'}
    path = None
    try:
        path = document.PathName
    except Exception:
        pass
    name = ProjectNameFromPath(path)
    cloud_path = None
    try:
        cloud_path = document.GetCloudModelPath()
    except Exception:
        pass
    if name is None and cloud_path is not None and db is not None:
        try:
            name = ProjectNameFromPath(db.ModelPathUtils.ConvertModelPathToUserVisiblePath(cloud_path))
        except Exception:
            pass
    try:
        hub_id = str(document.GetHubId() or '').strip()
    except Exception:
        hub_id = ''
    try:
        model_id = str(uuid.UUID(str(cloud_path.GetModelGUID())))
    except Exception:
        model_id = None
    return {'available': True, 'formaProjectId': project_id, 'formaProjectName': name,
            'formaHubId': hub_id, 'cloudModelGuid': model_id, 'nameSource': 'cloudPath' if name else None}


def CheckFormaProject(client, context, linked_project_id=None):
    """Lookup only. A failed lookup must never be treated as a missing project.

    linked_project_id is a Loadwise project the document was previously bound to.
    """
    found = client.find_forma_project({'formaProjectId': context['formaProjectId']})['project']
    linked_project = None
    if linked_project_id:
        linked_project = client.get_project({'projectId': str(uuid.UUID(linked_project_id))})['project']
    if found and linked_project and found['projectId'] != linked_project['projectId']:
        raise ValueError('This Revit model is linked to a different Loadwise project than its Autodesk Forma project. '
                         'No project link was changed. Review the existing connection before syncing.')
    return {'project': found, 'existingProject': None if found else linked_project, 'forma': dict(context)}


def CreateFormaProject(client, plan, name):
    """Call only after an explicit Create/Connect choice; the server prevents duplicates."""
    name = name.strip()
    if not name or len(name) > 200:
        raise ValueError('Enter a project name of 1 to 200 characters.')
    context = plan['forma']
    request = {'formaProjectId': context['formaProjectId'], 'formaHubId': context.get('formaHubId', ''), 'name': name}
    if plan.get('existingProject'):
        request['projectId'] = plan['existingProject']['projectId']
    return client.ensure_forma_project(request)
