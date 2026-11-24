-- Perbaikan keanggotaan (aman dijalankan berulang). Ganti email di baris INSERT bila perlu.
insert into public.members (email, name) values ('betterandcoid@gmail.com', 'Dea') on conflict (email) do nothing;
update public.members set email = lower(trim(email));
delete from public.members where email like '%@contoh.com';

grant select on public.members to authenticated;
grant select, insert, update, delete on public.projects, public.project_versions to authenticated;
grant usage, select on sequence public.project_versions_id_seq to authenticated;

create or replace function public.is_member() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.members m where lower(trim(m.email)) = lower(trim(coalesce(auth.jwt() ->> 'email', ''))));
$$;
drop policy if exists "member sees own row" on public.members;
create policy "member sees own row" on public.members for select to authenticated
  using (lower(trim(email)) = lower(trim(coalesce(auth.jwt() ->> 'email', ''))));

-- Cek hasil: harus muncul email lw
select email from public.members;
