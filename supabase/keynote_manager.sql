-- FFE Keynote Manager Supabase schema
-- Run this file in the Supabase SQL editor for the project used by pyRevit.

create table if not exists public.keynote_libraries (
  id uuid primary key default gen_random_uuid(),
  library_key text not null unique,
  display_path text not null default '',
  encoding text not null default 'utf-8',
  line_ending text not null default E'\r\n',
  file_hash text not null default '',
  last_write_utc double precision,
  dataset_version bigint not null default 1,
  entry_count integer not null default 0,
  last_saved_by_client_id text not null default '',
  last_saved_by_client_name text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.keynote_libraries
  add column if not exists file_hash text not null default '',
  add column if not exists last_write_utc double precision;

create table if not exists public.keynote_entries (
  id uuid primary key default gen_random_uuid(),
  library_id uuid not null references public.keynote_libraries(id) on delete cascade,
  keynote_key text not null,
  keynote_text text not null,
  parent_key text not null default '',
  sort_order integer not null default 0,
  row_version bigint not null default 1,
  updated_by_client_id text not null default '',
  updated_by_client_name text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint keynote_entries_key_not_empty check (btrim(keynote_key) <> ''),
  constraint keynote_entries_text_not_empty check (btrim(keynote_text) <> ''),
  constraint keynote_entries_key_no_tabs check (position(E'\t' in keynote_key) = 0 and position(E'\n' in keynote_key) = 0 and position(E'\r' in keynote_key) = 0),
  constraint keynote_entries_text_no_tabs check (position(E'\t' in keynote_text) = 0 and position(E'\n' in keynote_text) = 0 and position(E'\r' in keynote_text) = 0),
  constraint keynote_entries_parent_no_tabs check (position(E'\t' in parent_key) = 0 and position(E'\n' in parent_key) = 0 and position(E'\r' in parent_key) = 0),
  constraint keynote_entries_library_key_unique unique (library_id, keynote_key)
);

create index if not exists keynote_entries_library_order_idx
  on public.keynote_entries (library_id, sort_order, keynote_key);

create index if not exists keynote_entries_library_parent_idx
  on public.keynote_entries (library_id, parent_key);

create table if not exists public.keynote_edit_claims (
  id uuid primary key default gen_random_uuid(),
  library_id uuid not null references public.keynote_libraries(id) on delete cascade,
  claim_key text not null,
  db_id uuid,
  keynote_key text not null default '',
  client_id text not null,
  client_name text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint keynote_edit_claims_claim_key_not_empty check (btrim(claim_key) <> ''),
  constraint keynote_edit_claims_client_id_not_empty check (btrim(client_id) <> ''),
  constraint keynote_edit_claims_library_claim_unique unique (library_id, claim_key)
);

create index if not exists keynote_edit_claims_library_client_idx
  on public.keynote_edit_claims (library_id, client_id);

create or replace function public.touch_keynote_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists touch_keynote_libraries_updated_at on public.keynote_libraries;
create trigger touch_keynote_libraries_updated_at
before update on public.keynote_libraries
for each row execute function public.touch_keynote_updated_at();

drop trigger if exists touch_keynote_entries_updated_at on public.keynote_entries;
create trigger touch_keynote_entries_updated_at
before update on public.keynote_entries
for each row execute function public.touch_keynote_updated_at();

drop trigger if exists touch_keynote_edit_claims_updated_at on public.keynote_edit_claims;
create trigger touch_keynote_edit_claims_updated_at
before update on public.keynote_edit_claims
for each row execute function public.touch_keynote_updated_at();

create or replace function public.build_keynote_snapshot(p_library_id uuid)
returns jsonb
language sql
stable
security definer
set search_path = public
as $$
  select jsonb_build_object(
    'libraryId', l.id::text,
    'libraryKey', l.library_key,
    'displayPath', l.display_path,
    'encoding', l.encoding,
    'lineEnding', l.line_ending,
    'fileHash', l.file_hash,
    'lastWriteUtc', l.last_write_utc,
    'datasetVersion', l.dataset_version,
    'entryCount', l.entry_count,
    'updatedAt', l.updated_at,
    'lastSavedByClientId', l.last_saved_by_client_id,
    'lastSavedByClientName', l.last_saved_by_client_name,
    'entries', coalesce((
      select jsonb_agg(
        jsonb_build_object(
          'id', e.id::text,
          'dbId', e.id::text,
          'key', e.keynote_key,
          'text', e.keynote_text,
          'parentKey', e.parent_key,
          'sortOrder', e.sort_order,
          'rowVersion', e.row_version,
          'updatedAt', e.updated_at,
          'updatedByClientId', e.updated_by_client_id,
          'updatedByClientName', e.updated_by_client_name
        )
        order by e.sort_order, e.keynote_key
      )
      from public.keynote_entries e
      where e.library_id = l.id
    ), '[]'::jsonb)
  )
  from public.keynote_libraries l
  where l.id = p_library_id;
$$;

create or replace function public.validate_keynote_library(p_library_id uuid)
returns text[]
language plpgsql
stable
security definer
set search_path = public
as $$
declare
  v_errors text[] := array[]::text[];
  v_cycle record;
begin
  select coalesce(array_agg('Parent key "' || e.parent_key || '" was not found for key "' || e.keynote_key || '".'), array[]::text[])
  into v_errors
  from public.keynote_entries e
  where e.library_id = p_library_id
    and e.parent_key <> ''
    and not exists (
      select 1
      from public.keynote_entries p
      where p.library_id = e.library_id
        and p.keynote_key = e.parent_key
    );

  for v_cycle in
    with recursive walk(root_key, current_key, parent_key, path, has_cycle) as (
      select e.keynote_key, e.keynote_key, e.parent_key, array[e.keynote_key], false
      from public.keynote_entries e
      where e.library_id = p_library_id
      union all
      select w.root_key, p.keynote_key, p.parent_key, w.path || p.keynote_key, p.keynote_key = any(w.path)
      from walk w
      join public.keynote_entries p
        on p.library_id = p_library_id
       and p.keynote_key = w.parent_key
      where w.parent_key <> ''
        and not w.has_cycle
        and array_length(w.path, 1) < 100
    )
    select distinct root_key
    from walk
    where has_cycle
  loop
    v_errors := array_append(v_errors, 'Parent cycle detected at key "' || v_cycle.root_key || '".');
  end loop;

  return v_errors;
end;
$$;

create or replace function public.ensure_keynote_library(
  p_library_key text,
  p_display_path text,
  p_encoding text default 'utf-8',
  p_line_ending text default E'\r\n',
  p_seed_entries jsonb default '[]'::jsonb,
  p_client_id text default '',
  p_client_name text default ''
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_library_id uuid;
  v_entry record;
  v_errors text[];
begin
  if btrim(coalesce(p_library_key, '')) = '' then
    raise exception 'Library key is required.';
  end if;

  select id into v_library_id
  from public.keynote_libraries
  where library_key = p_library_key;

  if v_library_id is null then
    insert into public.keynote_libraries (
      library_key,
      display_path,
      encoding,
      line_ending,
      last_saved_by_client_id,
      last_saved_by_client_name
    )
    values (
      p_library_key,
      coalesce(p_display_path, ''),
      coalesce(nullif(p_encoding, ''), 'utf-8'),
      coalesce(nullif(p_line_ending, ''), E'\r\n'),
      coalesce(p_client_id, ''),
      coalesce(p_client_name, '')
    )
    returning id into v_library_id;

    for v_entry in
      select
        item.value ->> 'key' as key,
        item.value ->> 'text' as keynote_text,
        item.value ->> 'parentKey' as parent_key,
        item.ordinality
      from jsonb_array_elements(coalesce(p_seed_entries, '[]'::jsonb))
        with ordinality as item(value, ordinality)
    loop
      insert into public.keynote_entries (
        library_id,
        keynote_key,
        keynote_text,
        parent_key,
        sort_order,
        updated_by_client_id,
        updated_by_client_name
      )
      values (
        v_library_id,
        btrim(coalesce(v_entry.key, '')),
        btrim(coalesce(v_entry.keynote_text, '')),
        btrim(coalesce(v_entry.parent_key, '')),
        (v_entry.ordinality - 1)::integer,
        coalesce(p_client_id, ''),
        coalesce(p_client_name, '')
      );
    end loop;

    v_errors := public.validate_keynote_library(v_library_id);
    if array_length(v_errors, 1) is not null then
      raise exception 'Seed keynote data is invalid: %', array_to_string(v_errors, ' ');
    end if;

    update public.keynote_libraries
    set entry_count = (
          select count(*)
          from public.keynote_entries
          where library_id = v_library_id
        )
    where id = v_library_id;
  end if;

  return public.build_keynote_snapshot(v_library_id)
    || jsonb_build_object('status', 'ready', 'message', 'Loaded keynote library from Supabase.');
end;
$$;

create or replace function public.get_keynote_snapshot(p_library_key text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_library_id uuid;
begin
  select id into v_library_id
  from public.keynote_libraries
  where library_key = p_library_key;

  if v_library_id is null then
    return jsonb_build_object(
      'status', 'error',
      'message', 'Keynote library was not found.',
      'entries', '[]'::jsonb
    );
  end if;

  return public.build_keynote_snapshot(v_library_id)
    || jsonb_build_object('status', 'ready', 'message', 'Loaded keynote library from Supabase.');
end;
$$;

create or replace function public.build_keynote_edit_claims(p_library_id uuid)
returns jsonb
language sql
stable
security definer
set search_path = public
as $$
  select jsonb_build_object(
    'libraryId', l.id::text,
    'libraryKey', l.library_key,
    'claims', coalesce((
      select jsonb_agg(
        jsonb_build_object(
          'claimKey', c.claim_key,
          'dbId', coalesce(c.db_id::text, ''),
          'key', c.keynote_key,
          'clientId', c.client_id,
          'clientName', c.client_name,
          'updatedAt', c.updated_at
        )
        order by c.updated_at desc, c.claim_key
      )
      from public.keynote_edit_claims c
      where c.library_id = l.id
    ), '[]'::jsonb)
  )
  from public.keynote_libraries l
  where l.id = p_library_id;
$$;

create or replace function public.get_keynote_edit_claims(p_library_key text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_library_id uuid;
begin
  select id into v_library_id
  from public.keynote_libraries
  where library_key = p_library_key;

  if v_library_id is null then
    return jsonb_build_object(
      'status', 'error',
      'message', 'Keynote library was not found.',
      'claims', '[]'::jsonb
    );
  end if;

  return public.build_keynote_edit_claims(v_library_id)
    || jsonb_build_object('status', 'ready', 'message', 'Loaded keynote edit claims.');
end;
$$;

create or replace function public.replace_keynote_edit_claims(
  p_library_key text,
  p_client_id text default '',
  p_client_name text default '',
  p_claims jsonb default '[]'::jsonb
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_library_id uuid;
  v_claim record;
begin
  if btrim(coalesce(p_client_id, '')) = '' then
    raise exception 'Client id is required.';
  end if;

  select id into v_library_id
  from public.keynote_libraries
  where library_key = p_library_key;

  if v_library_id is null then
    return jsonb_build_object(
      'status', 'error',
      'message', 'Keynote library was not found.',
      'claims', '[]'::jsonb
    );
  end if;

  delete from public.keynote_edit_claims
  where library_id = v_library_id
    and client_id = p_client_id;

  for v_claim in
    select *
    from jsonb_to_recordset(coalesce(p_claims, '[]'::jsonb))
      as claim("claimKey" text, "dbId" text, key text)
  loop
    if btrim(coalesce(v_claim."claimKey", '')) <> '' then
      insert into public.keynote_edit_claims (
        library_id,
        claim_key,
        db_id,
        keynote_key,
        client_id,
        client_name
      )
      values (
        v_library_id,
        btrim(coalesce(v_claim."claimKey", '')),
        nullif(v_claim."dbId", '')::uuid,
        btrim(coalesce(v_claim.key, '')),
        p_client_id,
        coalesce(p_client_name, '')
      )
      on conflict (library_id, claim_key) do update
      set db_id = excluded.db_id,
          keynote_key = excluded.keynote_key,
          client_name = excluded.client_name
      where public.keynote_edit_claims.client_id = excluded.client_id;
    end if;
  end loop;

  return public.build_keynote_edit_claims(v_library_id)
    || jsonb_build_object('status', 'ready', 'message', 'Updated keynote edit claims.');
end;
$$;

create or replace function public.sync_keynote_file_snapshot(
  p_library_key text,
  p_display_path text,
  p_encoding text default 'utf-8',
  p_line_ending text default E'\r\n',
  p_file_hash text default '',
  p_last_write_utc double precision default null,
  p_entries jsonb default '[]'::jsonb,
  p_client_id text default '',
  p_client_name text default ''
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_library_id uuid;
  v_existing public.keynote_libraries%rowtype;
  v_entry record;
  v_errors text[];
begin
  if btrim(coalesce(p_library_key, '')) = '' then
    raise exception 'Library key is required.';
  end if;

  select *
  into v_existing
  from public.keynote_libraries
  where library_key = p_library_key;

  if found and coalesce(p_file_hash, '') <> '' and v_existing.file_hash = coalesce(p_file_hash, '') then
    return public.build_keynote_snapshot(v_existing.id)
      || jsonb_build_object('status', 'ready', 'message', 'Shared keynote file snapshot is already mirrored.');
  end if;

  insert into public.keynote_libraries (
    library_key,
    display_path,
    encoding,
    line_ending,
    file_hash,
    last_write_utc,
    last_saved_by_client_id,
    last_saved_by_client_name
  )
  values (
    p_library_key,
    coalesce(p_display_path, ''),
    coalesce(nullif(p_encoding, ''), 'utf-8'),
    coalesce(nullif(p_line_ending, ''), E'\r\n'),
    coalesce(p_file_hash, ''),
    p_last_write_utc,
    coalesce(p_client_id, ''),
    coalesce(p_client_name, '')
  )
  on conflict (library_key) do update
  set display_path = excluded.display_path,
      encoding = excluded.encoding,
      line_ending = excluded.line_ending,
      file_hash = excluded.file_hash,
      last_write_utc = excluded.last_write_utc,
      last_saved_by_client_id = excluded.last_saved_by_client_id,
      last_saved_by_client_name = excluded.last_saved_by_client_name
  returning id into v_library_id;

  delete from public.keynote_entries
  where library_id = v_library_id;

  for v_entry in
    select
      item.value ->> 'key' as key,
      item.value ->> 'text' as keynote_text,
      item.value ->> 'parentKey' as parent_key,
      item.ordinality
    from jsonb_array_elements(coalesce(p_entries, '[]'::jsonb))
      with ordinality as item(value, ordinality)
  loop
    insert into public.keynote_entries (
      library_id,
      keynote_key,
      keynote_text,
      parent_key,
      sort_order,
      updated_by_client_id,
      updated_by_client_name
    )
    values (
      v_library_id,
      btrim(coalesce(v_entry.key, '')),
      btrim(coalesce(v_entry.keynote_text, '')),
      btrim(coalesce(v_entry.parent_key, '')),
      (v_entry.ordinality - 1)::integer,
      coalesce(p_client_id, ''),
      coalesce(p_client_name, '')
    );
  end loop;

  v_errors := public.validate_keynote_library(v_library_id);
  if array_length(v_errors, 1) is not null then
    raise exception 'File keynote data is invalid: %', array_to_string(v_errors, ' ');
  end if;

  update public.keynote_libraries
  set dataset_version = dataset_version + 1,
      entry_count = (
        select count(*)
        from public.keynote_entries
        where library_id = v_library_id
      ),
      last_saved_by_client_id = coalesce(p_client_id, ''),
      last_saved_by_client_name = coalesce(p_client_name, '')
  where id = v_library_id;

  return public.build_keynote_snapshot(v_library_id)
    || jsonb_build_object('status', 'ready', 'message', 'Mirrored shared keynote file snapshot to Supabase.');
end;
$$;

create or replace function public.save_keynote_changes(
  p_library_key text,
  p_client_id text default '',
  p_client_name text default '',
  p_base_dataset_version bigint default 0,
  p_changes jsonb default '{}'::jsonb
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_library public.keynote_libraries%rowtype;
  v_item record;
  v_row public.keynote_entries%rowtype;
  v_metadata jsonb := coalesce(p_changes -> 'metadata', '{}'::jsonb);
  v_last_write_utc double precision;
  v_conflicts jsonb := '[]'::jsonb;
  v_conflict_count integer := 0;
  v_new_count integer := 0;
  v_touched_count integer := 0;
  v_deleted_count integer := 0;
  v_errors text[];
begin
  if nullif(v_metadata ->> 'lastWriteUtc', '') is not null then
    v_last_write_utc := (v_metadata ->> 'lastWriteUtc')::double precision;
  end if;

  select *
  into v_library
  from public.keynote_libraries
  where library_key = p_library_key
  for update;

  if not found then
    return jsonb_build_object(
      'status', 'error',
      'message', 'Keynote library was not found.',
      'entries', '[]'::jsonb
    );
  end if;

  for v_item in
    select *
    from jsonb_to_recordset(coalesce(p_changes -> 'deletes', '[]'::jsonb))
      as item("dbId" text, key text, "baseVersion" bigint)
  loop
    select *
    into v_row
    from public.keynote_entries
    where library_id = v_library.id
      and (
        (coalesce(v_item."dbId", '') <> '' and id = v_item."dbId"::uuid)
        or (coalesce(v_item."dbId", '') = '' and keynote_key = coalesce(v_item.key, ''))
      )
    for update;

    if not found and coalesce(v_item."dbId", '') <> '' then
      v_conflicts := v_conflicts || jsonb_build_array(jsonb_build_object(
        'type', 'delete',
        'key', coalesce(v_item.key, ''),
        'message', 'This keynote was already deleted.'
      ));
      v_conflict_count := v_conflict_count + 1;
    elsif coalesce(v_item."baseVersion", 0) > 0 and v_row.row_version <> v_item."baseVersion" then
      v_conflicts := v_conflicts || jsonb_build_array(jsonb_build_object(
        'type', 'delete',
        'key', v_row.keynote_key,
        'dbId', v_row.id::text,
        'rowVersion', v_row.row_version,
        'message', 'This keynote changed before it could be deleted.'
      ));
      v_conflict_count := v_conflict_count + 1;
    end if;
  end loop;

  for v_item in
    select *
    from jsonb_to_recordset(coalesce(p_changes -> 'upserts', '[]'::jsonb))
      as item("dbId" text, key text, "text" text, "parentKey" text, "sortOrder" integer, "baseVersion" bigint, "previousKey" text)
  loop
    select *
    into v_row
    from public.keynote_entries
    where library_id = v_library.id
      and (
        (coalesce(v_item."dbId", '') <> '' and id = v_item."dbId"::uuid)
        or (
          coalesce(v_item."dbId", '') = ''
          and coalesce(v_item."previousKey", '') <> ''
          and keynote_key = coalesce(v_item."previousKey", '')
        )
        or (
          coalesce(v_item."dbId", '') = ''
          and coalesce(v_item."previousKey", '') = ''
          and keynote_key = coalesce(v_item.key, '')
        )
      )
    for update;

    if found then
      if coalesce(v_item."baseVersion", 0) > 0 and v_row.row_version <> v_item."baseVersion" then
        v_conflicts := v_conflicts || jsonb_build_array(jsonb_build_object(
          'type', 'update',
          'key', v_row.keynote_key,
          'dbId', v_row.id::text,
          'rowVersion', v_row.row_version,
          'message', 'This keynote changed before your save.'
        ));
        v_conflict_count := v_conflict_count + 1;
      elsif exists (
        select 1
        from public.keynote_entries e
        where e.library_id = v_library.id
          and e.keynote_key = btrim(coalesce(v_item.key, ''))
          and e.id <> v_row.id
      ) then
        v_conflicts := v_conflicts || jsonb_build_array(jsonb_build_object(
          'type', 'duplicate',
          'key', coalesce(v_item.key, ''),
          'dbId', v_row.id::text,
          'message', 'Another keynote already uses this key.'
        ));
        v_conflict_count := v_conflict_count + 1;
      end if;
    elsif coalesce(v_item."dbId", '') <> '' then
        v_conflicts := v_conflicts || jsonb_build_array(jsonb_build_object(
          'type', 'update',
          'key', coalesce(v_item.key, ''),
          'dbId', v_item."dbId",
          'message', 'This keynote no longer exists.'
        ));
        v_conflict_count := v_conflict_count + 1;
    elsif exists (
      select 1
      from public.keynote_entries e
      where e.library_id = v_library.id
        and e.keynote_key = btrim(coalesce(v_item.key, ''))
    ) then
        v_conflicts := v_conflicts || jsonb_build_array(jsonb_build_object(
          'type', 'insert',
          'key', coalesce(v_item.key, ''),
          'message', 'Another keynote already uses this key.'
        ));
        v_conflict_count := v_conflict_count + 1;
    end if;
  end loop;

  if v_conflict_count > 0 then
    return jsonb_build_object(
      'status', 'conflict',
      'message', 'Some keynotes changed in Supabase before your save. Refresh or update the conflicting rows.',
      'conflicts', v_conflicts,
      'snapshot', public.build_keynote_snapshot(v_library.id)
    );
  end if;

  for v_item in
    select *
    from jsonb_to_recordset(coalesce(p_changes -> 'deletes', '[]'::jsonb))
      as item("dbId" text, key text, "baseVersion" bigint)
  loop
    delete from public.keynote_entries
    where library_id = v_library.id
      and (
        (coalesce(v_item."dbId", '') <> '' and id = v_item."dbId"::uuid)
        or (coalesce(v_item."dbId", '') = '' and keynote_key = coalesce(v_item.key, ''))
      );
    if found then
      v_deleted_count := v_deleted_count + 1;
    end if;
  end loop;

  for v_item in
    select *
    from jsonb_to_recordset(coalesce(p_changes -> 'upserts', '[]'::jsonb))
      as item("dbId" text, key text, "text" text, "parentKey" text, "sortOrder" integer, "baseVersion" bigint, "previousKey" text)
  loop
    update public.keynote_entries
    set keynote_key = btrim(coalesce(v_item.key, '')),
        keynote_text = btrim(coalesce(v_item."text", '')),
        parent_key = btrim(coalesce(v_item."parentKey", '')),
        sort_order = coalesce(v_item."sortOrder", sort_order),
        row_version = row_version + 1,
        updated_by_client_id = coalesce(p_client_id, ''),
        updated_by_client_name = coalesce(p_client_name, '')
    where library_id = v_library.id
      and (
        (coalesce(v_item."dbId", '') <> '' and id = v_item."dbId"::uuid)
        or (
          coalesce(v_item."dbId", '') = ''
          and coalesce(v_item."previousKey", '') <> ''
          and keynote_key = coalesce(v_item."previousKey", '')
        )
        or (
          coalesce(v_item."dbId", '') = ''
          and coalesce(v_item."previousKey", '') = ''
          and keynote_key = coalesce(v_item.key, '')
        )
      );

    if found then
      v_touched_count := v_touched_count + 1;
    else
      insert into public.keynote_entries (
        library_id,
        keynote_key,
        keynote_text,
        parent_key,
        sort_order,
        updated_by_client_id,
        updated_by_client_name
      )
      values (
        v_library.id,
        btrim(coalesce(v_item.key, '')),
        btrim(coalesce(v_item."text", '')),
        btrim(coalesce(v_item."parentKey", '')),
        coalesce(v_item."sortOrder", 0),
        coalesce(p_client_id, ''),
        coalesce(p_client_name, '')
      );
      v_new_count := v_new_count + 1;
    end if;
  end loop;

  v_errors := public.validate_keynote_library(v_library.id);
  if array_length(v_errors, 1) is not null then
    raise exception 'Saved keynote data is invalid: %', array_to_string(v_errors, ' ');
  end if;

  update public.keynote_libraries
  set dataset_version = dataset_version + 1,
      entry_count = (
        select count(*)
        from public.keynote_entries
        where library_id = v_library.id
      ),
      display_path = coalesce(nullif(v_metadata ->> 'displayPath', ''), nullif(display_path, ''), p_library_key),
      encoding = coalesce(nullif(v_metadata ->> 'encoding', ''), encoding),
      line_ending = coalesce(nullif(v_metadata ->> 'lineEnding', ''), line_ending),
      file_hash = coalesce(v_metadata ->> 'fileHash', file_hash),
      last_write_utc = coalesce(v_last_write_utc, last_write_utc),
      last_saved_by_client_id = coalesce(p_client_id, ''),
      last_saved_by_client_name = coalesce(p_client_name, '')
  where id = v_library.id;

  return public.build_keynote_snapshot(v_library.id)
    || jsonb_build_object(
      'status', 'ready',
      'message', 'Saved keynote changes to Supabase.',
      'insertedCount', v_new_count,
      'updatedCount', v_touched_count,
      'deletedCount', v_deleted_count,
      'baseDatasetVersion', p_base_dataset_version
    );
end;
$$;

alter table public.keynote_libraries enable row level security;
alter table public.keynote_entries enable row level security;
alter table public.keynote_edit_claims enable row level security;

drop policy if exists "anon can read keynote libraries" on public.keynote_libraries;
create policy "anon can read keynote libraries"
on public.keynote_libraries
for select
to anon, authenticated
using (true);

drop policy if exists "anon can read keynote entries" on public.keynote_entries;
create policy "anon can read keynote entries"
on public.keynote_entries
for select
to anon, authenticated
using (true);

drop policy if exists "anon can read keynote edit claims" on public.keynote_edit_claims;
create policy "anon can read keynote edit claims"
on public.keynote_edit_claims
for select
to anon, authenticated
using (true);

revoke all on public.keynote_libraries from anon, authenticated;
revoke all on public.keynote_entries from anon, authenticated;
revoke all on public.keynote_edit_claims from anon, authenticated;
grant select on public.keynote_libraries to anon, authenticated;
grant select on public.keynote_entries to anon, authenticated;
grant select on public.keynote_edit_claims to anon, authenticated;

revoke execute on function public.build_keynote_snapshot(uuid) from public;
revoke execute on function public.build_keynote_edit_claims(uuid) from public;
revoke execute on function public.validate_keynote_library(uuid) from public;
revoke execute on function public.ensure_keynote_library(text, text, text, text, jsonb, text, text) from public;
revoke execute on function public.get_keynote_snapshot(text) from public;
revoke execute on function public.get_keynote_edit_claims(text) from public;
revoke execute on function public.replace_keynote_edit_claims(text, text, text, jsonb) from public;
revoke execute on function public.sync_keynote_file_snapshot(text, text, text, text, text, double precision, jsonb, text, text) from public;
revoke execute on function public.save_keynote_changes(text, text, text, bigint, jsonb) from public;

grant execute on function public.ensure_keynote_library(text, text, text, text, jsonb, text, text) to anon, authenticated;
grant execute on function public.get_keynote_snapshot(text) to anon, authenticated;
grant execute on function public.get_keynote_edit_claims(text) to anon, authenticated;
grant execute on function public.replace_keynote_edit_claims(text, text, text, jsonb) to anon, authenticated;
grant execute on function public.sync_keynote_file_snapshot(text, text, text, text, text, double precision, jsonb, text, text) to anon, authenticated;
grant execute on function public.save_keynote_changes(text, text, text, bigint, jsonb) to anon, authenticated;

do $$
begin
  if not exists (
    select 1
    from pg_publication_tables
    where pubname = 'supabase_realtime'
      and schemaname = 'public'
      and tablename = 'keynote_libraries'
  ) then
    alter publication supabase_realtime add table public.keynote_libraries;
  end if;

  if not exists (
    select 1
    from pg_publication_tables
    where pubname = 'supabase_realtime'
      and schemaname = 'public'
      and tablename = 'keynote_edit_claims'
  ) then
    alter publication supabase_realtime add table public.keynote_edit_claims;
  end if;
end;
$$;
