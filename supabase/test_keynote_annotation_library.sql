-- Run after keynote_annotation_library.sql. All fixtures are rolled back.
begin;
set local role anon;
do $$
declare
  v_key text := 'annotation:test-' || gen_random_uuid()::text;
  v_snapshot jsonb;
  v_result jsonb;
  v_version bigint;
  v_id text;
  v_file_key text := 'test-export-' || gen_random_uuid()::text || '.txt';
  v_library_id text;
begin
  v_snapshot := public.ensure_annotation_keynote_library(v_key, 'Test library');
  assert jsonb_array_length(v_snapshot->'entries') = 21, 'Template divisions';
  v_version := (v_snapshot->>'datasetVersion')::bigint;
  v_id := v_snapshot->'entries'->0->>'dbId';
  v_result := public.save_annotation_keynote_changes(v_key, 'test-client', 'Test', v_version,
    jsonb_build_object('upserts', jsonb_build_array(jsonb_build_object(
      'dbId', v_id, 'key', 'DIVISION 01', 'text', '', 'parentKey', '', 'baseVersion', 1))));
  assert v_result->>'status' = 'ready', 'Direct cloud save';
  assert v_result->'entries'->0->>'text' = '', 'Blank descriptions preserved';
  v_result := public.save_annotation_keynote_changes(v_key, 'other-client', 'Other', v_version, '{}');
  assert v_result->>'status' = 'conflict', 'Stale database version rejected';
  v_result := public.ensure_annotation_keynote_library(v_key, 'Test', '[]');
  assert jsonb_array_length(v_result->'entries') = 21, 'Existing cloud library must not be reseeded';
  assert v_result->'entries'->0->>'text' = '', 'Canonical cloud edit must survive setup';
  begin
    perform public.sync_keynote_file_snapshot(v_key, 'not-a-file', 'utf-8', E'\r\n', 'bad-hash', null, '[]');
    raise exception 'TEST FAILED: legacy mirror overwrote cloud library';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%annotation save API%', 'Legacy mirror explicitly rejected';
  end;
  begin
    update public.keynote_templates set content = 'changed' where template_key = 'ffe-divisions';
    raise exception 'TEST FAILED: template is editable by app clients';
  exception when insufficient_privilege then null;
  end;

  v_snapshot := public.get_keynote_snapshot(v_key);
  v_library_id := v_snapshot->>'libraryId';
  v_version := (v_snapshot->>'datasetVersion')::bigint;
  begin
    perform public.convert_annotation_keynote_library_to_file(v_key, v_version - 1,
      v_file_key, v_file_key, 'hash', 123);
    raise exception 'TEST FAILED: stale export converted library';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%changed during export%', 'Concurrent edits block conversion';
  end;
  v_result := public.convert_annotation_keynote_library_to_file(v_key, v_version,
    v_file_key, v_file_key, 'hash', 123);
  assert v_result->>'libraryId' = v_library_id, 'Conversion keeps original library ID';
  assert v_result->>'sourceType' = 'file', 'Conversion changes source type';
  assert v_result->>'libraryKey' = v_file_key, 'Conversion updates path identity';
  assert v_result->>'fileHash' = 'hash', 'Conversion stores export metadata';
  assert v_result->'entries' = v_snapshot->'entries', 'Conversion preserves every row and row ID';
  assert (public.get_keynote_snapshot(v_key))->>'status' = 'error', 'No duplicate annotation record';
  assert (public.get_annotation_keynote_snapshot(v_key))->>'libraryId' = v_library_id, 'Project alias survives conversion';
  v_result := public.convert_annotation_keynote_library_to_file(v_key, v_version,
    v_file_key, v_file_key, 'hash', 123);
  assert v_result->>'libraryId' = v_library_id, 'Retry is idempotent';
  v_result := public.ensure_annotation_keynote_library(v_key, 'Test', '[]');
  assert v_result->>'libraryId' = v_library_id, 'Switching back reuses the same record';
  assert v_result->'entries' = v_snapshot->'entries', 'Switching back retains all rows';
  raise notice 'Annotation library SQL acceptance checks passed';
end;
$$;
rollback;
