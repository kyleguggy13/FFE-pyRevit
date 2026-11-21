# -*- coding: utf-8 -*-
"""Autodesk Forma project discovery and startup connection, with injectable APIs.

ReadFormaContext is called only in a valid Revit API context. Everything returned
is plain JSON; CheckFormaProject/CreateFormaProject can then run off-thread.
No credentials are obtained from Revit, and no Space data or model metadata is written.
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


def ModelIdForDocument(linked_id, context):
    """Keep existing bindings; unlinked cloud copies use the same model UUID."""
    if linked_id:
        return str(uuid.UUID(linked_id))
    if context and context.get('available') and context.get('cloudModelGuid'):
        return str(uuid.uuid5(uuid.NAMESPACE_URL, 'ffe-mechanical:forma:{0}:{1}'.format(
            context['formaProjectId'], context['cloudModelGuid'])))
    return str(uuid.uuid4())


def _pages(operation, request):
    page = operation(request)
    while True:
        for item in page['items']:
            yield item
        if page['nextCursor'] is None:
            break
        page = operation(dict(request, cursor=page['nextCursor']))


def CheckFormaProject(client, context, linked_model_id=None):
    """Lookup only. A failed lookup must never be treated as a missing project."""
    found = client.find_forma_project({'formaProjectId': context['formaProjectId']})['project']
    linked_project = None
    if linked_model_id:
        for project in _pages(client.list_projects, {'limit': 200}):
            models = _pages(client.list_models, {'projectId': project['projectId'], 'limit': 200})
            if any(model['modelId'] == linked_model_id for model in models):
                linked_project = project
                break
    if found and linked_project and found['projectId'] != linked_project['projectId']:
        raise ValueError('This Revit model is linked to a different repository than its Autodesk Forma project. '
                         'No project or model link was changed. Review the existing connection before publishing.')
    return {'project': found, 'existingProject': linked_project, 'forma': dict(context)}


def CreateFormaProject(client, plan, name, request_id):
    """Call only after an explicit Create/Connect choice; the server prevents duplicates."""
    name = name.strip()
    if not name or len(name) > 200:
        raise ValueError('Enter a project name of 1 to 200 characters.')
    context = plan['forma']
    request = {'formaProjectId': context['formaProjectId'], 'formaHubId': context.get('formaHubId', ''),
               'name': name, 'requestId': request_id}
    if plan.get('existingProject'):
        request['projectId'] = plan['existingProject']['projectId']
    return client.ensure_forma_project(request)
