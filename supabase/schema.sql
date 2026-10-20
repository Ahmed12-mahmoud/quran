-- شغّل هذا الملف مرة واحدة في Supabase SQL Editor.
-- بعد إنشاء مستخدم المشرف من Authentication > Users، عطّل التسجيل العام.

create table if not exists public.question_edits (
  uid text primary key,
  text text not null,
  updated_at timestamptz not null default now(),
  updated_by uuid not null references auth.users(id) default auth.uid()
);

alter table public.question_edits enable row level security;

revoke all on table public.question_edits from anon;
grant select, insert, update, delete on table public.question_edits to authenticated;

drop policy if exists "admins can read edits" on public.question_edits;
drop policy if exists "admins can insert edits" on public.question_edits;
drop policy if exists "admins can update edits" on public.question_edits;
drop policy if exists "admins can delete edits" on public.question_edits;

-- لا يسمح بالقراءة أو التعديل إلا للمستخدمين المسجلين.
-- عطّل Signups من Authentication > Providers بعد إنشاء حسابك حتى يبقى الحساب لك وحدك.
create policy "admins can read edits"
  on public.question_edits for select to authenticated
  using (auth.uid() is not null);

create policy "admins can insert edits"
  on public.question_edits for insert to authenticated
  with check (auth.uid() = updated_by);

create policy "admins can update edits"
  on public.question_edits for update to authenticated
  using (auth.uid() = updated_by)
  with check (auth.uid() = updated_by);

create policy "admins can delete edits"
  on public.question_edits for delete to authenticated
  using (auth.uid() = updated_by);
