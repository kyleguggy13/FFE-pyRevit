-- Apply after keynote_manager.sql, before installing the Generic Annotation Only client.
-- Additive migration: existing file libraries keep their API and data.
begin;

create table if not exists public.keynote_templates (
  template_key text primary key,
  version integer not null check (version > 0),
  content text not null,
  updated_at timestamptz not null default now()
);
alter table public.keynote_templates enable row level security;
drop policy if exists "clients can read keynote templates" on public.keynote_templates;
create policy "clients can read keynote templates" on public.keynote_templates
  for select to anon, authenticated using (true);
revoke all on public.keynote_templates from anon, authenticated;
grant select on public.keynote_templates to anon, authenticated;

insert into public.keynote_templates(template_key, version, content)
values ('ffe-divisions', 1, replace(replace($template$# @datastore("txt") @source("\\172.16.1.7\ffe\Internal Share\Drafting\Revit\Templates\Keynotes\RevitKeynotes_TEMPLATE.txt") @encoding("UTF-16")
# --------------------- @table(categories:"Root Keynotes Table")
DIVISION 01	GENERAL REQUIREMENTS
DIVISION 02	EXISTING CONDITIONS
DIVISION 03	CONCRETE
DIVISION 04	MASONRY
DIVISION 05	METALS
DIVISION 06	WOOD, PLASTICS AND COMPOSITES
DIVISION 07	THERMAL AND MOISTURE PROTECTION
DIVISION 08	OPENINGS
DIVISION 09	FINISHES
DIVISION 10	SPECIALTIES
DIVISION 11	EQUIPMENT
DIVISION 12	FURNISHINGS
DIVISION 13	SPECIAL CONSTRUCTION
DIVISION 14	CONVEYING EQUIPMENT
DIVISION 22	PLUMBING
DIVISION 23	HEATING, VENTILATING AND AIR CONDITIONING
DIVISION 26	ELECTRICAL
DIVISION 28	ELECTRONIC SAFETY AND SECURITY
DIVISION 31	EARTHWORK
DIVISION 32	EXTERIOR IMPROVEMENTS
DIVISION 33	UTILITIES
# --------------------- @table(keynotes:"Keynotes Table")
$template$, E'\r\n', E'\n'), E'\n', E'\r\n'))
on conflict (template_key) do nothing;

alter table public.keynote_libraries
  add column if not exists source_type text not null default 'file';
alter table public.keynote_libraries
  add column if not exists annotation_library_key text;
create unique index if not exists keynote_libraries_annotation_origin
  on public.keynote_libraries(annotation_library_key) where annotation_library_key is not null;
update public.keynote_libraries set annotation_library_key = library_key
  where source_type = 'annotation' and annotation_library_key is null;
-- Blank descriptions are valid in the existing text-file syntax.
alter table public.keynote_entries drop constraint if exists keynote_entries_text_not_empty;

-- Keep the existing snapshot contract, adding the selected source metadata.
create or replace function public.get_keynote_snapshot(p_library_key text)
returns jsonb language plpgsql security definer set search_path = public
set statement_timeout = '60s'
as $$
declare v_library public.keynote_libraries%rowtype;
begin
  select * into v_library from public.keynote_libraries where library_key = p_library_key;
  if not found then
    return jsonb_build_object('status', 'error', 'message', 'Keynote library was not found.', 'entries', '[]'::jsonb);
  end if;
  return public.build_keynote_snapshot(v_library.id) || jsonb_build_object(
    'status', 'ready', 'message', 'Loaded keynote library from Supabase.', 'sourceType', v_library.source_type);
end;
$$;

create or replace function public.get_keynote_template(p_template_key text default 'ffe-divisions')
returns jsonb language sql stable security definer set search_path = public
as $$
  select jsonb_build_object('templateKey', template_key, 'version', version, 'content', content)
  from public.keynote_templates where template_key = p_template_key;
$$;


-- Remove obsolete draft RVT-mirror APIs if that draft was previously installed.
drop trigger if exists guard_model_keynote_entries on public.keynote_entries;
drop trigger if exists guard_model_keynote_library on public.keynote_libraries;
drop function if exists public.guard_model_keynote_write();
drop function if exists public.get_model_keynote_snapshot(text);
drop function if exists public.sync_model_keynote_snapshot(text,text,jsonb,jsonb,text,text,text);

create or replace function public.ensure_annotation_keynote_library(
  p_library_key text, p_display_path text, p_seed_entries jsonb default null
)
returns jsonb language plpgsql security definer set search_path = public
as $$
declare
  v_library public.keynote_libraries%rowtype;
  v_content text;
  v_seed jsonb := p_seed_entries;
begin
  if coalesce(p_library_key, '') not like 'annotation:%' then
    raise exception 'Annotation library identity is required.';
  end if;
  perform pg_advisory_xact_lock(hashtextextended(p_library_key, 0));
  select * into v_library from public.keynote_libraries
    where library_key = p_library_key or annotation_library_key = p_library_key for update;
  if found then
    if v_library.source_type = 'file' and v_library.annotation_library_key = p_library_key then
      update public.keynote_libraries set source_type = 'annotation',
        dataset_version = dataset_version + 1, display_path = p_display_path
        where id = v_library.id returning * into v_library;
    elsif v_library.source_type <> 'annotation' then
      raise exception 'Library source mismatch.';
    end if;
    return public.get_keynote_snapshot(v_library.library_key);
  end if;
  if v_seed is null then
    select content into v_content from public.keynote_templates where template_key = 'ffe-divisions';
    if v_content is null then raise exception 'The FFE division template has not been uploaded.'; end if;
    select coalesce(jsonb_agg(jsonb_build_object(
      'key', split_part(line, E'\t', 1), 'text', split_part(line, E'\t', 2),
      'parentKey', split_part(line, E'\t', 3)) order by ordinal), '[]'::jsonb)
    into v_seed
    from regexp_split_to_table(v_content, E'\r?\n') with ordinality as lines(line, ordinal)
    where btrim(line) <> '' and ltrim(line) not like '#%';
    if jsonb_array_length(v_seed) = 0 then raise exception 'The FFE template is empty.'; end if;
  end if;
  perform public.ensure_keynote_library(p_library_key, p_display_path, 'utf-8', E'\r\n', v_seed);
  update public.keynote_libraries set source_type = 'annotation', annotation_library_key = p_library_key
    where library_key = p_library_key;
  return public.get_keynote_snapshot(p_library_key);
end;
$$;

-- Protect canonical cloud libraries from legacy file-mirror clients.
create or replace function public.guard_annotation_keynote_write()
returns trigger language plpgsql set search_path = public
as $$
declare v_key text; v_type text;
begin
  if TG_TABLE_NAME = 'keynote_entries' then
    select library_key, source_type into v_key, v_type from public.keynote_libraries
    where id = case when TG_OP = 'DELETE' then OLD.library_id else NEW.library_id end;
  else
    v_key := OLD.library_key;
    v_type := OLD.source_type;
    if NEW.library_key = OLD.library_key and NEW.source_type = OLD.source_type and NEW.file_hash = OLD.file_hash
       and NEW.dataset_version = OLD.dataset_version and NEW.entry_count = OLD.entry_count then
      return NEW;
    end if;
  end if;
  if v_type = 'annotation' and current_setting('ffe.annotation_write', true) is distinct from v_key then
    raise exception 'Annotation libraries require the Supabase annotation save API.';
  end if;
  if TG_OP = 'DELETE' then return OLD; end if;
  return NEW;
end;
$$;
drop trigger if exists guard_annotation_keynote_entries on public.keynote_entries;
create trigger guard_annotation_keynote_entries before insert or update or delete on public.keynote_entries
for each row execute function public.guard_annotation_keynote_write();
drop trigger if exists guard_annotation_keynote_library on public.keynote_libraries;
create trigger guard_annotation_keynote_library before update on public.keynote_libraries
for each row execute function public.guard_annotation_keynote_write();

create or replace function public.save_annotation_keynote_changes(
  p_library_key text, p_client_id text default '', p_client_name text default '',
  p_base_dataset_version bigint default 0, p_changes jsonb default '{}'::jsonb
)
returns jsonb language plpgsql security definer set search_path = public
set statement_timeout = '60s'
as $$
declare v_library public.keynote_libraries%rowtype; v_result jsonb;
begin
  select * into v_library from public.keynote_libraries where library_key = p_library_key for update;
  if not found or v_library.source_type <> 'annotation' then
    raise exception 'Annotation library was not found.';
  end if;
  if v_library.dataset_version <> p_base_dataset_version then
    return jsonb_build_object('status','conflict', 'message','The Supabase library changed. Refresh before saving.',
      'snapshot', public.get_keynote_snapshot(p_library_key));
  end if;
  -- Existing edit claims remain advisory in file mode; enforce touched claims here.
  if exists (
    select 1 from public.keynote_edit_claims c
    join jsonb_array_elements(coalesce(p_changes->'upserts','[]') || coalesce(p_changes->'deletes','[]')) changes(value) on
      c.keynote_key in (changes.value->>'key', changes.value->>'previousKey') or c.db_id::text = changes.value->>'dbId'
    where c.library_id = v_library.id and c.client_id <> coalesce(p_client_id,'')
      and c.updated_at > now() - interval '3 minutes'
  ) then
    return jsonb_build_object('status','conflict', 'message','Another user is editing an affected keynote. Retry after they release it.');
  end if;
  perform set_config('ffe.annotation_write', p_library_key, true);
  v_result := public.save_keynote_changes(p_library_key, p_client_id, p_client_name, p_base_dataset_version, p_changes);
  perform set_config('ffe.annotation_write', '', true);
  if v_result->>'status' <> 'ready' then return v_result; end if;
  return public.get_keynote_snapshot(p_library_key);
end;
$$;

revoke all on function public.get_keynote_template(text) from public;
revoke all on function public.ensure_annotation_keynote_library(text,text,jsonb) from public;
revoke all on function public.save_annotation_keynote_changes(text,text,text,bigint,jsonb) from public;
grant execute on function public.get_keynote_template(text) to anon, authenticated;
grant execute on function public.ensure_annotation_keynote_library(text,text,jsonb) to anon, authenticated;
grant execute on function public.save_annotation_keynote_changes(text,text,text,bigint,jsonb) to anon, authenticated;

create or replace function public.get_annotation_keynote_snapshot(p_library_key text)
returns jsonb language sql stable security definer set search_path = public
as $$
  select public.get_keynote_snapshot(coalesce(
    (select library_key from public.keynote_libraries
      where library_key = p_library_key or annotation_library_key = p_library_key limit 1), p_library_key));
$$;

-- Move the existing library to file mode. Preserve library/entry IDs, claims, and analytics.
create or replace function public.convert_annotation_keynote_library_to_file(
  p_annotation_key text, p_base_dataset_version bigint, p_file_key text, p_display_path text,
  p_file_hash text, p_last_write_utc double precision, p_encoding text default 'utf-16',
  p_line_ending text default E'\r\n', p_client_id text default '', p_client_name text default ''
)
returns jsonb language plpgsql security definer set search_path = public
as $$
declare v_library public.keynote_libraries%rowtype;
begin
  if coalesce(p_annotation_key, '') not like 'annotation:%' or coalesce(p_file_key, '') = ''
     or p_file_key like 'annotation:%' or coalesce(p_file_hash, '') = '' then
    raise exception 'The annotation identity, destination file, and file hash are required.';
  end if;
  perform pg_advisory_xact_lock(hashtextextended(p_annotation_key, 0));
  select * into v_library from public.keynote_libraries
    where library_key = p_annotation_key or annotation_library_key = p_annotation_key for update;
  if not found then raise exception 'The original annotation library was not found.'; end if;
  -- Safe retry if the first response was lost after the conversion committed.
  if v_library.source_type = 'file' and v_library.library_key = p_file_key and v_library.file_hash = p_file_hash then
    return public.get_keynote_snapshot(p_file_key);
  end if;
  if p_base_dataset_version is null or v_library.dataset_version <> p_base_dataset_version then
    raise exception 'The library changed during export. Retry to export the latest notes.';
  end if;
  if exists (select 1 from public.keynote_libraries where library_key = p_file_key and id <> v_library.id) then
    raise exception 'That file already belongs to another Supabase library. Choose a different filename.';
  end if;
  perform set_config('ffe.annotation_write', v_library.library_key, true);
  update public.keynote_libraries set
    library_key = p_file_key, annotation_library_key = p_annotation_key, source_type = 'file',
    display_path = p_display_path, encoding = p_encoding, line_ending = p_line_ending,
    file_hash = p_file_hash, last_write_utc = p_last_write_utc, dataset_version = dataset_version + 1,
    last_saved_by_client_id = p_client_id, last_saved_by_client_name = p_client_name
    where id = v_library.id;
  perform set_config('ffe.annotation_write', '', true);
  return public.get_keynote_snapshot(p_file_key);
end;
$$;
revoke all on function public.get_annotation_keynote_snapshot(text) from public;
revoke all on function public.convert_annotation_keynote_library_to_file(text,bigint,text,text,text,double precision,text,text,text,text) from public;
grant execute on function public.get_annotation_keynote_snapshot(text) to anon, authenticated;
grant execute on function public.convert_annotation_keynote_library_to_file(text,bigint,text,text,text,double precision,text,text,text,text) to anon, authenticated;

-- Promote a file mirror in place. Retain the library ID and existing row IDs.
create or replace function public.convert_file_keynote_library_to_annotation(
  p_file_key text, p_annotation_key text, p_base_dataset_version bigint,
  p_display_path text, p_file_hash text, p_entries jsonb,
  p_client_id text default '', p_client_name text default ''
)
returns jsonb language plpgsql security definer set search_path = public
set statement_timeout = '60s'
as $$
declare v_library public.keynote_libraries%rowtype; v_errors text[];
begin
  if coalesce(p_file_key, '') = '' or p_file_key like 'annotation:%'
     or coalesce(p_annotation_key, '') not like 'annotation:%'
     or coalesce(p_file_hash, '') = '' or p_entries is null or jsonb_typeof(p_entries) <> 'array' then
    raise exception 'The file identity, annotation identity, hash, and entries are required.';
  end if;
  perform pg_advisory_xact_lock(hashtextextended(p_annotation_key, 0));
  select * into v_library from public.keynote_libraries where library_key = p_file_key for update;
  if not found then
    raise exception 'The original file library was not found. Refresh before retrying.';
  end if;
  -- A lost conversion response must never reimport the old file over cloud edits.
  if v_library.source_type = 'annotation' and v_library.annotation_library_key = p_annotation_key then
    return public.get_keynote_snapshot(v_library.library_key);
  end if;
  if v_library.source_type <> 'file' then
    raise exception 'This library is already used by another annotation project.';
  end if;
  if p_base_dataset_version is null or v_library.dataset_version <> p_base_dataset_version then
    raise exception 'The library changed during conversion. Refresh before retrying.';
  end if;
  if exists (select 1 from public.keynote_libraries where id <> v_library.id
             and (library_key = p_annotation_key or annotation_library_key = p_annotation_key)) then
    raise exception 'This project already has a different Supabase library. Conversion would replace its association.';
  end if;
  if v_library.annotation_library_key is not null and v_library.annotation_library_key <> p_annotation_key then
    raise exception 'This file library is associated with another project.';
  end if;
  -- The assigned file is authoritative until this transaction switches the source.
  -- Import external file edits without replacing rows whose keys still exist.
  if v_library.file_hash <> p_file_hash then
    insert into public.keynote_entries (library_id, keynote_key, keynote_text, parent_key, sort_order,
                                       updated_by_client_id, updated_by_client_name)
    select v_library.id, btrim(coalesce(item.value->>'key', '')), btrim(coalesce(item.value->>'text', '')),
           btrim(coalesce(item.value->>'parentKey', '')), (item.ordinality - 1)::integer,
           coalesce(p_client_id, ''), coalesce(p_client_name, '')
      from jsonb_array_elements(p_entries) with ordinality as item(value, ordinality)
    on conflict (library_id, keynote_key) do update set
      keynote_text = excluded.keynote_text, parent_key = excluded.parent_key, sort_order = excluded.sort_order,
      row_version = keynote_entries.row_version + 1,
      updated_by_client_id = excluded.updated_by_client_id, updated_by_client_name = excluded.updated_by_client_name
    where (keynote_entries.keynote_text, keynote_entries.parent_key, keynote_entries.sort_order)
       is distinct from (excluded.keynote_text, excluded.parent_key, excluded.sort_order);
    delete from public.keynote_entries where library_id = v_library.id
      and keynote_key not in (select btrim(coalesce(value->>'key', '')) from jsonb_array_elements(p_entries));
    v_errors := public.validate_keynote_library(v_library.id);
    if array_length(v_errors, 1) is not null then
      raise exception 'File keynote data is invalid: %', array_to_string(v_errors, ' ');
    end if;
  end if;
  -- Keep the file key too: legacy mirror clients then find this protected record
  -- instead of creating a second record under the original path.
  update public.keynote_libraries set annotation_library_key = p_annotation_key,
    source_type = 'annotation', display_path = p_display_path, file_hash = p_file_hash,
    dataset_version = dataset_version + 1,
    entry_count = (select count(*) from public.keynote_entries where library_id = v_library.id),
    last_saved_by_client_id = coalesce(p_client_id, ''), last_saved_by_client_name = coalesce(p_client_name, '')
    where id = v_library.id;
  return public.get_keynote_snapshot(v_library.library_key);
end;
$$;
revoke all on function public.convert_file_keynote_library_to_annotation(text,text,bigint,text,text,jsonb,text,text) from public;
grant execute on function public.convert_file_keynote_library_to_annotation(text,text,bigint,text,text,jsonb,text,text) to anon, authenticated;
commit;
