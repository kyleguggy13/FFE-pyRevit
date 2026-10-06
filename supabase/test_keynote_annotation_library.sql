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

-- First conversion of an ordinary file mirror, including retry and external edits.
do $$
declare
  v_file_key text := 'test-file-' || gen_random_uuid()::text || '.txt';
  v_key text := 'annotation:test-file-project-' || gen_random_uuid()::text;
  v_export_key text := 'test-reexport-' || gen_random_uuid()::text || '.txt';
  v_entries jsonb := '[{"key":"DIVISION 22","text":"PLUMBING"},
                        {"key":"22.01","text":"","parentKey":"DIVISION 22"},
                        {"key":"22.01A","text":"Café","parentKey":"22.01"}]';
  v_snapshot jsonb; v_result jsonb; v_claims jsonb; v_analytics jsonb;
  v_library_id text; v_version bigint; v_other_key text; v_other_file text;
begin
  v_result := public.sync_keynote_file_snapshot(v_file_key, v_file_key, 'utf-16', E'\r\n', 'original-hash', 123, v_entries);
  v_snapshot := public.get_keynote_snapshot(v_file_key);
  v_library_id := v_snapshot->>'libraryId';
  v_version := (v_snapshot->>'datasetVersion')::bigint;
  perform public.replace_keynote_edit_claims(v_file_key, 'claim-client', 'Claim user',
    jsonb_build_array(jsonb_build_object('claimKey', 'key:22.01', 'key', '22.01',
                                       'dbId', v_snapshot->'entries'->1->>'dbId')));
  v_claims := (public.get_keynote_edit_claims(v_file_key))->'claims';
  perform public.sync_keynote_analytics(v_file_key, 'test-model', 'Test model');
  select to_jsonb(d) into v_analytics from public.keynote_analytics_documents d
    where library_id::text = v_library_id and document_key = 'test-model';
  v_snapshot := public.get_keynote_snapshot(v_file_key);
  begin
    perform public.convert_file_keynote_library_to_annotation(v_file_key, v_key, v_version - 1,
      'Test project', 'original-hash', v_entries);
    raise exception 'TEST FAILED: stale import converted library';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%changed during conversion%', 'Stale file conversion rejected';
  end;
  begin
    perform public.convert_file_keynote_library_to_annotation(v_file_key, v_key, v_version,
      'Test project', 'invalid-hash', '[{"key":"bad","text":"bad","parentKey":"missing"}]');
    raise exception 'TEST FAILED: malformed import converted library';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%File keynote data is invalid%', 'Invalid import rejected';
  end;
  assert public.get_keynote_snapshot(v_file_key) = v_snapshot, 'Failed conversion rolls back imported rows and metadata';
  v_result := public.convert_file_keynote_library_to_annotation(v_file_key, v_key, v_version,
    'Test project', 'original-hash', v_entries);
  assert v_result->>'libraryId' = v_library_id, 'First file-to-annotation switch preserves library ID';
  assert v_result->>'libraryKey' = v_file_key, 'File identity retained to prevent duplicate mirrors';
  assert v_result->>'sourceType' = 'annotation', 'Library source changes in place';
  assert v_result->'entries' = v_snapshot->'entries', 'All entries, IDs, and versions retained';
  assert (select count(*) from public.keynote_libraries where library_key in (v_file_key, v_key)
          or annotation_library_key = v_key) = 1, 'Only one library record exists';
  assert (public.get_annotation_keynote_snapshot(v_key))->>'libraryId' = v_library_id, 'Project alias resolves converted file';
  assert (public.get_keynote_edit_claims(v_file_key))->'claims' = v_claims, 'Claims retained';
  assert (select to_jsonb(d) from public.keynote_analytics_documents d
          where library_id::text = v_library_id and document_key = 'test-model') = v_analytics, 'Analytics retained';
  v_result := public.ensure_annotation_keynote_library(v_key, 'Test project', '[]');
  assert v_result->>'libraryId' = v_library_id, 'Repeated setup resolves alias without reseeding';
  assert v_result->'entries' = v_snapshot->'entries', 'Repeated setup preserves notes';
  -- Cloud edits after conversion must survive a retry with the original file contents/version.
  v_result := public.save_annotation_keynote_changes(v_file_key, 'claim-client', 'Claim user',
    (v_result->>'datasetVersion')::bigint, jsonb_build_object('upserts', jsonb_build_array(jsonb_build_object(
      'dbId', v_snapshot->'entries'->1->>'dbId', 'key', '22.01', 'text', 'Cloud edit',
      'parentKey', 'DIVISION 22', 'baseVersion', 1))));
  v_snapshot := v_result;
  v_result := public.convert_file_keynote_library_to_annotation(v_file_key, v_key, v_version,
    'Test project', 'original-hash', v_entries);
  assert v_result = v_snapshot, 'Retry retains newer cloud changes';
  begin
    perform public.sync_keynote_file_snapshot(v_file_key, v_file_key, 'utf-16', E'\r\n', 'stale-client-hash', 123, v_entries);
    raise exception 'TEST FAILED: old file client overwrote converted library';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%annotation save API%', 'Legacy file mirror cannot overwrite converted record';
  end;
  -- Convert to another filename, then import edits made externally without replacing surviving row IDs.
  v_result := public.convert_annotation_keynote_library_to_file(v_key,
    (v_snapshot->>'datasetVersion')::bigint, v_export_key, v_export_key, 'export-hash', 123);
  v_version := (v_result->>'datasetVersion')::bigint;
  v_result := public.convert_file_keynote_library_to_annotation(v_export_key, v_key, v_version,
    'Test project', 'external-edit-hash', '[{"key":"DIVISION 22","text":"PLUMBING"},
      {"key":"22.01","text":"External edit","parentKey":"DIVISION 22"},
      {"key":"22.02","text":"New note","parentKey":"DIVISION 22"}]');
  assert v_result->>'libraryId' = v_library_id, 'Round trip retains original library';
  assert v_result->'entries'->0 = v_snapshot->'entries'->0, 'Unchanged file row retains all metadata';
  assert v_result->'entries'->1->>'dbId' = v_snapshot->'entries'->1->>'dbId', 'Externally edited row keeps its ID';
  assert v_result->'entries'->1->>'text' = 'External edit', 'Latest file text imported';
  assert v_result->'entries'->2->>'key' = '22.02', 'External additions/deletions imported';
  assert (v_result->'entries'->1->>'rowVersion')::bigint =
         (v_snapshot->'entries'->1->>'rowVersion')::bigint + 1, 'Edited row version advances';
  assert (public.get_annotation_keynote_snapshot(v_key))->>'libraryId' = v_library_id, 'Alias survives round trip';
  -- A separate existing project library must not be replaced by conversion.
  v_other_key := 'annotation:test-existing-' || gen_random_uuid()::text;
  v_other_file := 'test-collision-' || gen_random_uuid()::text || '.txt';
  perform public.ensure_annotation_keynote_library(v_other_key, 'Existing project', v_entries);
  v_result := public.sync_keynote_file_snapshot(v_other_file, v_other_file, 'utf-16', E'\r\n', 'hash', 123, v_entries);
  begin
    perform public.convert_file_keynote_library_to_annotation(v_other_file, v_other_key,
      (v_result->>'datasetVersion')::bigint, 'Test', 'hash', v_entries);
    raise exception 'TEST FAILED: existing project association replaced';
  exception when others then
    if SQLERRM like 'TEST FAILED:%' then raise; end if;
    assert SQLERRM like '%different Supabase library%', 'Existing project association protected';
  end;
  assert (public.get_keynote_snapshot(v_other_file))->>'sourceType' = 'file', 'Conflicting conversion leaves source unchanged';
  raise notice 'File-to-annotation SQL acceptance checks passed';
end;
$$;
rollback;
