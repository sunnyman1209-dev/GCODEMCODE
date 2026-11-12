-- G·M코드 퀴즈 응시 기록 표
-- Supabase 화면 왼쪽 'SQL Editor'에 통째로 붙여 넣고 Run을 누른다. 한 번만 실행하면 된다.

create table if not exists public.gm_quiz_records (
  id         bigint generated always as identity primary key,
  created_at timestamptz not null default now(),          -- 응시 시각(서버가 찍음)
  name       text not null check (char_length(name) between 1 and 20),
  score      int  not null check (score between 0 and 100),
  passed     boolean not null,
  browser    text check (char_length(browser) <= 60),
  check (passed = (score >= 90))
);

-- 학생(공개 키)은 기록을 '추가'만 할 수 있다. 읽기·고치기·지우기는 못 한다.
alter table public.gm_quiz_records enable row level security;
drop policy if exists "students insert only" on public.gm_quiz_records;
create policy "students insert only" on public.gm_quiz_records
  for insert to anon with check (true);

-- 응시 시각은 학생이 보낼 수 없고 서버 시각으로만 찍힌다.
revoke all on public.gm_quiz_records from anon, authenticated;
grant insert (name, score, passed, browser) on public.gm_quiz_records to anon;
