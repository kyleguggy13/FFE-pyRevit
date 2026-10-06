-- Apply after keynote_manager.sql and keynote_annotation_library.sql.
-- IDs remain canonical; old and new paths resolve to the same record.
begin;
alter table public.keynote_libraries add column if not exists file_library_key text;
update public.keynote_libraries set file_library_key = library_key
  where source_type = 'file' and file_library_key is null;

create or replace function public.track_keynote_library_file_identity()
returns trigger language plpgsql set search_path = public
as $$
begin
  if NEW.source_type = 'file' and (TG_OP = 'INSERT' or OLD.source_type <> 'file' or NEW.library_key <> OLD.library_key) then
    NEW.file_library_key := NEW.library_key;
  end if;
  return NEW;
end;
$$;
drop trigger if exists track_keynote_library_file_identity on public.keynote_libraries;
create trigger track_keynote_library_file_identity before insert or update of library_key, source_type
on public.keynote_libraries for each row execute function public.track_keynote_library_file_identity();

create or replace function public.guard_keynote_active_file()
returns trigger language plpgsql set search_path = public
as $$
begin
  if OLD.source_type = 'file' and NEW.source_type = 'file'
     and (NEW.file_hash is distinct from OLD.file_hash or NEW.dataset_version > OLD.dataset_version)
     and lower(replace(NEW.display_path, E'\\', '/')) <>
         lower(replace(coalesce(OLD.file_library_key, OLD.library_key), E'\\', '/')) then
    raise exception 'The keynote file moved. Reconnect to the library current file before saving or mirroring.';
  end if;
  return NEW;
end;
$$;
drop trigger if exists guard_keynote_active_file on public.keynote_libraries;
create trigger guard_keynote_active_file before update on public.keynote_libraries
for each row execute function public.guard_keynote_active_file();

create table if not exists public.keynote_library_aliases (
  alias_key text primary key check (btrim(alias_key) <> ''),
  library_id uuid not null references public.keynote_libraries(id) on delete cascade
);
alter table public.keynote_library_aliases enable row level security;
revoke all on public.keynote_library_aliases from public, anon, authenticated;

create or replace function public.register_keynote_library_aliases()
returns trigger language plpgsql set search_path = public
as $$
declare v_key text;
begin
  foreach v_key in array array[NEW.library_key, NEW.annotation_library_key] loop
    if coalesce(v_key, '') = '' then continue; end if;
    insert into public.keynote_library_aliases(alias_key, library_id) values (v_key, NEW.id)
      on conflict (alias_key) do nothing;
    if not exists (select 1 from public.keynote_library_aliases where alias_key = v_key and library_id = NEW.id) then
      raise exception 'That keynote identity already belongs to another library. Reconnect or choose another file.';
    end if;
  end loop;
  return NEW;
end;
$$;
-- Reject conflicting legacy identities rather than silently assigning their data.
do $$
begin
  if exists (select alias_key from (
    select library_key as alias_key, id from public.keynote_libraries
    union select annotation_library_key, id from public.keynote_libraries where annotation_library_key is not null
  ) origins group by alias_key having count(distinct id) > 1) then
    raise exception 'Conflicting keynote library identities exist. Resolve them before installing stable library identity.';
  end if;
end;
$$;
insert into public.keynote_library_aliases(alias_key, library_id)
select library_key, id from public.keynote_libraries
union select annotation_library_key, id from public.keynote_libraries where annotation_library_key is not null
on conflict (alias_key) do nothing;
drop trigger if exists register_keynote_library_aliases on public.keynote_libraries;
create trigger register_keynote_library_aliases after insert or update of library_key, annotation_library_key
on public.keynote_libraries for each row execute function public.register_keynote_library_aliases();

create or replace function public.get_keynote_snapshot(p_library_key text)
returns jsonb language plpgsql security definer set search_path = public
set statement_timeout = '60s'
as $$
declare v_library public.keynote_libraries%rowtype;
begin
  select * into v_library from public.keynote_libraries
    where library_key = p_library_key or id =
      (select library_id from public.keynote_library_aliases where alias_key = p_library_key);
  if not found then
    return jsonb_build_object('status','error','message','Keynote library was not found.','entries','[]'::jsonb);
  end if;
  return public.build_keynote_snapshot(v_library.id) || jsonb_build_object(
    'status','ready','message','Loaded keynote library from Supabase.','sourceType',v_library.source_type,
    'annotationLibraryKey',v_library.annotation_library_key,
    'fileLibraryKey',coalesce(v_library.file_library_key, case when v_library.source_type = 'file' then v_library.library_key end));
end;
$$;

create or replace function public.get_keynote_snapshot_by_id(p_library_id uuid)
returns jsonb language plpgsql security definer set search_path = public
as $$
declare v_library public.keynote_libraries%rowtype;
begin
  select * into v_library from public.keynote_libraries where id = p_library_id;
  if not found then
    return jsonb_build_object('status','error','message','The associated keynote library was not found. Restore the library or explicitly choose another association.');
  end if;
  return public.get_keynote_snapshot(v_library.library_key);
end;
$$;

create or replace function public.link_keynote_library_path(p_library_id uuid, p_file_key text, p_display_path text)
returns jsonb language plpgsql security definer set search_path = public
as $$
declare v_library public.keynote_libraries%rowtype;
begin
  if btrim(coalesce(p_file_key,'')) = '' or p_file_key like 'annotation:%' then raise exception 'A file identity is required.'; end if;
  select * into v_library from public.keynote_libraries where id = p_library_id for update;
  if not found or v_library.source_type <> 'file' then raise exception 'The associated library is not in Text File mode. Refresh and review Storage Mode.'; end if;
  if exists (select 1 from public.keynote_libraries where library_key = p_file_key and id <> p_library_id)
     or exists (select 1 from public.keynote_library_aliases where alias_key = p_file_key and library_id <> p_library_id) then
    raise exception 'The selected file is associated with another keynote library. Neither association was changed.';
  end if;
  insert into public.keynote_library_aliases values (p_file_key, p_library_id) on conflict (alias_key) do nothing;
  if not exists (select 1 from public.keynote_library_aliases where alias_key = p_file_key and library_id = p_library_id) then
    raise exception 'The selected file was associated with another library during reconnection. Retry.';
  end if;
  if v_library.display_path is distinct from p_display_path or v_library.file_library_key is distinct from p_file_key then
    update public.keynote_libraries set display_path = p_display_path, file_library_key = p_file_key where id = p_library_id;
  end if;
  return public.get_keynote_snapshot_by_id(p_library_id);
end;
$$;

-- Updated clients resolve aliases before invoking the existing creation/mirror API.
create or replace function public.ensure_keynote_library_association(
  p_library_id uuid, p_library_key text, p_file_key text, p_display_path text,
  p_encoding text default 'utf-8', p_line_ending text default E'\r\n',
  p_seed_entries jsonb default '[]'::jsonb, p_client_id text default '', p_client_name text default ''
)
returns jsonb language plpgsql security definer set search_path = public
as $$
declare v_result jsonb;
begin
  if p_library_id is not null then
    v_result := public.get_keynote_snapshot_by_id(p_library_id);
    if v_result->>'status' <> 'ready' then return v_result; end if;
  else
    v_result := public.get_keynote_snapshot(p_library_key);
    if v_result->>'status' <> 'ready' then
      v_result := public.ensure_keynote_library(p_library_key, p_display_path, p_encoding, p_line_ending,
                                               p_seed_entries, p_client_id, p_client_name);
    end if;
  end if;
  if v_result->>'fileLibraryKey' is not null and v_result->>'fileLibraryKey' <> p_file_key then
    raise exception 'The keynote file moved. Reconnect to % before attaching its library.', v_result->>'displayPath';
  end if;
  return public.link_keynote_library_path((v_result->>'libraryId')::uuid, p_file_key, p_display_path);
end;
$$;

create or replace function public.fork_keynote_library(
  p_source_library_id uuid, p_base_dataset_version bigint, p_library_key text, p_display_path text,
  p_source_type text, p_seed_entries jsonb default null, p_file_hash text default '',
  p_last_write_utc double precision default null, p_encoding text default 'utf-8', p_line_ending text default E'\r\n'
)
returns jsonb language plpgsql security definer set search_path = public
as $$
declare v_source public.keynote_libraries%rowtype; v_entries jsonb; v_result jsonb; v_id uuid;
begin
  if p_source_type not in ('file','annotation') or btrim(coalesce(p_library_key,'')) = ''
     or (p_source_type = 'annotation') <> (p_library_key like 'annotation:%')
     or (p_source_type = 'file' and coalesce(p_file_hash,'') = '') then
    raise exception 'A valid destination identity and storage mode are required.';
  end if;
  select * into v_source from public.keynote_libraries where id = p_source_library_id for update;
  if not found or p_base_dataset_version is null or v_source.dataset_version <> p_base_dataset_version then
    raise exception 'The original library changed. Refresh before creating an independent library.';
  end if;
  if exists (select 1 from public.keynote_libraries where library_key = p_library_key)
     or exists (select 1 from public.keynote_library_aliases where alias_key = p_library_key) then
    raise exception 'The destination already belongs to a keynote library. Choose another destination.';
  end if;
  -- File contents are authoritative in file mode; annotation forks copy saved cloud notes.
  v_entries := case when v_source.source_type = 'file' then p_seed_entries else null end;
  if v_entries is null then
    select coalesce(jsonb_agg(jsonb_build_object('key',keynote_key,'text',keynote_text,'parentKey',parent_key)
           order by sort_order,keynote_key),'[]'::jsonb) into v_entries
      from public.keynote_entries where library_id = v_source.id;
  end if;
  v_result := public.ensure_keynote_library(p_library_key,p_display_path,p_encoding,p_line_ending,v_entries);
  v_id := (v_result->>'libraryId')::uuid;
  update public.keynote_libraries set source_type = p_source_type,
    annotation_library_key = case when p_source_type = 'annotation' then p_library_key else null end,
    file_hash = p_file_hash, last_write_utc = p_last_write_utc where id = v_id;
  return public.get_keynote_snapshot_by_id(v_id);
end;
$$;

create or replace function public.list_keynote_libraries()
returns jsonb language sql stable security definer set search_path = public
as $$
  select jsonb_build_object('status','ready','libraries',coalesce(jsonb_agg(jsonb_build_object(
    'libraryId',id::text,'libraryKey',library_key,'displayPath',display_path,'sourceType',source_type)
    order by display_path,library_key),'[]'::jsonb)) from public.keynote_libraries;
$$;

revoke all on function public.get_keynote_snapshot_by_id(uuid) from public;
revoke all on function public.link_keynote_library_path(uuid,text,text) from public;
revoke all on function public.ensure_keynote_library_association(uuid,text,text,text,text,text,jsonb,text,text) from public;
revoke all on function public.fork_keynote_library(uuid,bigint,text,text,text,jsonb,text,double precision,text,text) from public;
revoke all on function public.list_keynote_libraries() from public;
grant execute on function public.get_keynote_snapshot_by_id(uuid) to anon, authenticated;
grant execute on function public.link_keynote_library_path(uuid,text,text) to anon, authenticated;
grant execute on function public.ensure_keynote_library_association(uuid,text,text,text,text,text,jsonb,text,text) to anon, authenticated;
grant execute on function public.fork_keynote_library(uuid,bigint,text,text,text,jsonb,text,double precision,text,text) to anon, authenticated;
grant execute on function public.list_keynote_libraries() to anon, authenticated;
commit;
