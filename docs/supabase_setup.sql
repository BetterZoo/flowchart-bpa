-- Flowchart BPA v1.9.0 — Supabase setup
-- Run once in Supabase: SQL Editor → New query → paste everything → Run.
-- Then add every team member's email at the bottom (and re-run only that INSERT when adding people).

create table if not exists public.members (
  email      text primary key,
  name       text,
  added_at   timestamptz not null default now()
);

create table if not exists public.projects (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  file_name   text,
  data        jsonb not null default '{}'::jsonb,   -- snapshot of the app's project data
  version     integer not null default 1,           -- optimistic lock: a save only succeeds on the version it started from
  created_by  text,
  updated_by  text,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);

create table if not exists public.project_versions (
  id          bigserial primary key,
  project_id  uuid not null references public.projects(id) on delete cascade,
  version     integer,
  data        jsonb not null,
  note        text,
  saved_by    text,
  saved_at    timestamptz not null default now()
);
create index if not exists project_versions_project_idx on public.project_versions(project_id, id desc);

-- Who is a team member? (email of the signed-in user, case-insensitive)
create or replace function public.is_member() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.members m where lower(m.email) = lower(coalesce(auth.jwt() ->> 'email', '')));
$$;

alter table public.members          enable row level security;
alter table public.projects         enable row level security;
alter table public.project_versions enable row level security;

drop policy if exists "member sees own row"  on public.members;
drop policy if exists "team reads members"   on public.members;
create policy "member sees own row" on public.members for select to authenticated
  using (lower(email) = lower(coalesce(auth.jwt() ->> 'email', '')));

drop policy if exists "team projects" on public.projects;
create policy "team projects" on public.projects for all to authenticated
  using (public.is_member()) with check (public.is_member());

drop policy if exists "team versions" on public.project_versions;
create policy "team versions" on public.project_versions for all to authenticated
  using (public.is_member()) with check (public.is_member());

revoke all on public.members, public.projects, public.project_versions from anon;
grant select on public.members to authenticated;
grant select, insert, update, delete on public.projects, public.project_versions to authenticated;
grant usage, select on sequence public.project_versions_id_seq to authenticated;

-- Keep only the 30 newest history versions per project
create or replace function public.prune_project_versions() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  delete from public.project_versions
   where project_id = new.project_id
     and id not in (select id from public.project_versions where project_id = new.project_id order by id desc limit 30);
  return new;
end $$;
drop trigger if exists trg_prune_project_versions on public.project_versions;
create trigger trg_prune_project_versions after insert on public.project_versions
  for each row execute function public.prune_project_versions();

-- ▼ Team members (edit the emails, add one line per person)
insert into public.members (email, name) values
  ('GANTI_DENGAN_EMAIL_LW@contoh.com', 'Dea')
on conflict (email) do nothing;
