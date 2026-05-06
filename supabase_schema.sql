create table if not exists public.allowed_users (
  email text primary key,
  role text not null check (role in ('admin', 'moderator')),
  created_at timestamptz not null default now()
);

alter table public.allowed_users enable row level security;

create or replace function public.current_user_email()
returns text
language sql
stable
as $$
  select lower(auth.jwt() ->> 'email')
$$;

create or replace function public.is_allowed_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.allowed_users
    where email = public.current_user_email()
      and role = 'admin'
  )
$$;

drop policy if exists "allowed_users_select_own_or_admin" on public.allowed_users;
create policy "allowed_users_select_own_or_admin"
on public.allowed_users
for select
to authenticated
using (
  email = public.current_user_email()
  or public.is_allowed_admin()
);

drop policy if exists "allowed_users_insert_admin" on public.allowed_users;
create policy "allowed_users_insert_admin"
on public.allowed_users
for insert
to authenticated
with check (public.is_allowed_admin());

drop policy if exists "allowed_users_update_admin" on public.allowed_users;
create policy "allowed_users_update_admin"
on public.allowed_users
for update
to authenticated
using (public.is_allowed_admin())
with check (public.is_allowed_admin());

drop policy if exists "allowed_users_delete_admin" on public.allowed_users;
create policy "allowed_users_delete_admin"
on public.allowed_users
for delete
to authenticated
using (public.is_allowed_admin());

-- 初回だけ、Supabase SQL Editorで自分のGoogleアカウントをadminとして追加してください。
-- insert into public.allowed_users (email, role)
-- values ('your-email@example.com', 'admin')
-- on conflict (email) do update set role = excluded.role;
