-- Run after keynote_library_identity.sql. All fixtures are rolled back.
begin;
set local role anon;
do $$
declare
  v_key text := 'c:/ffe/identity-' || gen_random_uuid()::text || '.txt';
  v_moved text := 'c:/ffe/moved-' || gen_random_uuid()::text || '.txt';
  v_copy text := 'c:/ffe/copy-' || gen_random_uuid()::text || '.txt';
  v_other text := 'c:/ffe/other-' || gen_random_uuid()::text || '.txt';
  v_annotation text := 'annotation:identity-copy-' || gen_random_uuid()::text;
  v_second_annotation text := 'annotation:identity-fork-' || gen_random_uuid()::text;
  v_bad text := 'annotation:identity-stale-' || gen_random_uuid()::text;
  v_entries jsonb := '[{"key":"DIVISION 22","text":"PLUMBING"},
                       {"key":"22.01","text":"","parentKey":"DIVISION 22"},
                       {"key":"22.01A","text":"Café","parentKey":"22.01"}]';
  v_original jsonb; v_snapshot jsonb; v_result jsonb; v_id uuid; v_missing uuid := gen_random_uuid();
  v_version bigint; v_count bigint; v_fork_id uuid; v_claims jsonb; v_analytics jsonb;
begin
  -- Windows display paths and normalized keys may use different separators/case.
  perform public.sync_keynote_file_snapshot(v_key, replace(upper(v_key), '/', E'\\'), 'utf-16', E'\r\n', 'hash', 123, v_entries);
  v_original := public.get_keynote_snapshot(v_key);
  v_id := (v_original->>'libraryId')::uuid;
  v_version := (v_original->>'datasetVersion')::bigint;
  perform public.replace_keynote_edit_claims(v_key, 'identity-client', 'Test',
    jsonb_build_array(jsonb_build_object('claimKey','key:22.01','key','22.01','dbId',v_original->'entries'->1->>'dbId')));
  v_claims := (public.get_keynote_edit_claims(v_key))->'claims';
  perform public.sync_keynote_analytics(v_key, 'identity-model', 'Identity model');
  select to_jsonb(d) into v_analytics from public.keynote_analytics_documents d where library_id = v_id;
  v_snapshot := public.link_keynote_library_path(v_id, v_moved, replace(upper(v_moved), '/', E'\\'));
  assert v_snapshot->>'libraryId' = v_id::text, 'Move retains UUID';
  assert v_snapshot->>'fileLibraryKey' = v_moved, 'Move updates active file';
  assert v_snapshot->>'libraryKey' = v_key, 'Canonical key remains stable';
  assert (v_snapshot->>'datasetVersion')::bigint = v_version, 'Metadata move does not create a data revision';
  assert v_snapshot->'entries' = v_original->'entries', 'Move preserves note IDs and every row';
  assert (public.get_keynote_snapshot(v_key))->>'libraryId' = v_id::text, 'Old path remains an alias';
  assert (public.get_keynote_snapshot(v_moved))->>'libraryId' = v_id::text, 'New path resolves to same record';
  assert (public.get_keynote_snapshot_by_id(v_id))->>'libraryId' = v_id::text, 'UUID lookup ignores file name';
  assert (public.get_keynote_edit_claims(v_key))->'claims' = v_claims, 'Claims retained through move';
  assert (select to_jsonb(d) from public.keynote_analytics_documents d where library_id = v_id) = v_analytics, 'Analytics retained';
  v_result := public.ensure_keynote_library_association(v_id, v_key, v_moved, v_snapshot->>'displayPath');
  assert v_result->>'libraryId' = v_id::text, 'Bound client retains ID';
  v_result := public.ensure_keynote_library_association(null, v_moved, v_moved, v_snapshot->>'displayPath');
  assert v_result->>'libraryId' = v_id::text, 'Unbound client discovers path alias without duplicate';
  select count(*) into v_count from public.keynote_libraries;
  v_result := public.ensure_keynote_library_association(v_missing, 'missing-file.txt', 'missing-file.txt', 'missing-file.txt');
  assert v_result->>'status' = 'error', 'Missing UUID rejects attachment';
  assert (select count(*) from public.keynote_libraries) = v_count, 'Missing UUID never falls back to creating by path';
  begin
    perform public.sync_keynote_file_snapshot(v_key, v_key, 'utf-16', E'\r\n', 'stale-hash', 123, '[]');
    raise exception 'TEST FAILED: stale original path overwrote moved library';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%file moved%', 'Old active file cannot mirror or save';
  end;
  begin
    perform public.ensure_keynote_library(v_moved, v_moved, 'utf-16', E'\r\n', v_entries);
    raise exception 'TEST FAILED: legacy client created duplicate at alias path';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%already belongs%', 'Legacy creation cannot steal an alias';
  end;
  assert (select count(*) from public.keynote_libraries) = v_count, 'Failed legacy writes create no duplicates';
  assert (public.get_keynote_snapshot_by_id(v_id))->'entries' = v_original->'entries', 'Failed writes preserve notes';
  perform public.sync_keynote_file_snapshot(v_other, v_other, 'utf-16', E'\r\n', 'other-hash', 123, v_entries);
  begin
    perform public.link_keynote_library_path(v_id, v_other, v_other);
    raise exception 'TEST FAILED: move stole another library path';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%another keynote library%', 'Path collision preserves both associations';
  end;
  -- File forks use the current valid file content, with fresh row IDs.
  v_result := public.fork_keynote_library(v_id, v_version, v_copy, v_copy, 'file', v_entries, 'copy-hash', 123, 'utf-16');
  v_fork_id := (v_result->>'libraryId')::uuid;
  assert v_fork_id <> v_id, 'Independent file library has a new UUID';
  assert v_result->'entries'->1->>'dbId' <> v_original->'entries'->1->>'dbId', 'Independent note IDs differ';
  assert v_result->'entries'->1->>'text' = '', 'Blank descriptions preserved';
  assert v_result->'entries'->2->>'text' = 'Café', 'Unicode preserved';
  assert v_result->'entries'->2->>'parentKey' = '22.01', 'Nested hierarchy preserved';
  assert not exists (select 1 from public.keynote_edit_claims where library_id = v_fork_id), 'Claims are not copied';
  assert not exists (select 1 from public.keynote_analytics_documents where library_id = v_fork_id), 'Analytics are not copied';
  assert (public.get_keynote_snapshot_by_id(v_id))->'entries' = v_original->'entries', 'Fork leaves source notes intact';
  begin
    perform public.fork_keynote_library(v_id, v_version - 1, v_bad, 'Stale', 'annotation');
    raise exception 'TEST FAILED: stale fork succeeded';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%original library changed%', 'Fork checks source revision';
  end;
  assert (public.get_keynote_snapshot(v_bad))->>'status' = 'error', 'Failed fork creates no record';
  -- The UUID also resolves after a conversion changes the canonical key.
  v_result := public.convert_file_keynote_library_to_annotation(v_key, v_annotation, v_version, 'Test project', 'hash', v_entries);
  v_result := public.convert_annotation_keynote_library_to_file(v_annotation, (v_result->>'datasetVersion')::bigint,
    v_moved, v_moved, 'export-hash', 123);
  assert v_result->>'libraryId' = v_id::text, 'Conversion reuses UUID';
  assert (public.get_keynote_snapshot(v_key))->>'libraryId' = v_id::text, 'Historical canonical key stays an alias';
  assert (public.get_keynote_snapshot_by_id(v_id))->>'libraryKey' = v_moved, 'UUID resolves new canonical key';
  -- Annotation forks copy saved server notes, ignoring a supplied local snapshot.
  v_result := public.ensure_annotation_keynote_library(v_annotation, 'Test project');
  v_result := public.fork_keynote_library(v_id, (v_result->>'datasetVersion')::bigint, v_second_annotation, 'Copy',
    'annotation', '[{"key":"wrong","text":"unsaved local data"}]');
  assert v_result->>'libraryId' <> v_id::text, 'Independent annotation library has new UUID';
  assert jsonb_array_length(v_result->'entries') = 3, 'Annotation fork copies complete saved library';
  assert v_result->'entries'->2->>'text' = 'Café', 'Annotation fork preserves saved Unicode notes';
  assert exists (select 1 from jsonb_array_elements((public.list_keynote_libraries())->'libraries') as item
    where item->>'libraryId' = v_id::text), 'Recovery lists original library';
  raise notice 'Stable keynote library identity checks passed';
end;
$$;
rollback;
