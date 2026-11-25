-- Flowchart BPA v1.9.8 — log setiap simpan + snapshot saat ganti penyunting
-- Jalankan sekali di Supabase SQL Editor (aman dijalankan ulang).

-- 1) Satu baris per simpan (tanpa isi data): siapa, kapan, versi, bagian yang berubah
create table if not exists public.project_saves (
  id          bigserial primary key,
  project_id  uuid not null references public.projects(id) on delete cascade,
  version     integer,
  saved_by    text,
  saved_at    timestamptz not null default now(),
  changed     text[] not null default '{}'
);
create index if not exists project_saves_project_idx on public.project_saves(project_id, id desc);
alter table public.project_saves enable row level security;
drop policy if exists "team saves read" on public.project_saves;
create policy "team saves read" on public.project_saves for select to authenticated using (public.is_member());
revoke all on public.project_saves from anon;
grant select on public.project_saves to authenticated;

-- 2) Trigger di projects: catat setiap simpan; bila penyunting berganti, simpan dulu
--    versi terakhir penyunting sebelumnya ke Riwayat (bila belum ada) supaya pekerjaannya
--    tidak "hilang" di antara snapshot 15 menit.
create or replace function public.log_project_save() returns trigger
language plpgsql security definer set search_path = public as $$
declare ch text[];
begin
  if new.data is not distinct from old.data and new.version is not distinct from old.version then
    return new;
  end if;
  select coalesce(array_agg(replace(k, 'flowbpa:', '') order by k), '{}') into ch
    from (select key k from jsonb_each(new.data) union select key from jsonb_each(old.data)) ks
   where (new.data -> k) is distinct from (old.data -> k);
  insert into public.project_saves(project_id, version, saved_by, saved_at, changed)
       values (new.id, new.version, new.updated_by, coalesce(new.updated_at, now()), ch);
  if old.updated_by is distinct from new.updated_by
     and not exists (select 1 from public.project_versions v where v.project_id = old.id and v.version = old.version) then
    insert into public.project_versions(project_id, version, data, note, saved_by, saved_at)
         values (old.id, old.version, old.data, 'Versi terakhir sebelum diedit ' || coalesce(split_part(new.updated_by, '@', 1), 'orang lain'),
                 old.updated_by, coalesce(old.updated_at, now()));
  end if;
  -- keep the log bounded: 500 newest saves per project
  delete from public.project_saves where project_id = new.id
     and id < (select min(id) from (select id from public.project_saves where project_id = new.id order by id desc limit 500) t);
  return new;
end $$;
drop trigger if exists trg_log_project_save on public.projects;
create trigger trg_log_project_save after update on public.projects
  for each row execute function public.log_project_save();
