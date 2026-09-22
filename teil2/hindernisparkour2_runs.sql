-- Ranking-Backend für Teil 2 des Hindernisparcours.
-- Einmal im Supabase-Dashboard unter "SQL Editor" ausführen.
--
-- Aufbau bewusst identisch zu Teil 1, nur mit eigener Tabelle und eigenen
-- Funktionsnamen (hindernisparkour2_*), damit die Zeiten beider Teile nicht
-- vermischt werden. Die Tabelle selbst bleibt über die öffentliche API
-- gesperrt (RLS an, keine Policies) – geschrieben und gelesen wird
-- ausschließlich über die drei Funktionen hier.

create table if not exists public.hindernisparkour2_runs (
  anon_id    text primary key,
  seconds    double precision not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.hindernisparkour2_runs enable row level security;

create index if not exists hindernisparkour2_runs_seconds_idx
  on public.hindernisparkour2_runs (seconds);

-- Trägt einen Lauf ein. Pro Gerät (anon_id) bleibt nur die beste Zeit stehen,
-- beliebig viele Versuche zählen also nie als weitere Teilnehmer.
-- Zurück kommen Rang, Teilnehmerzahl, Bestzeit aller und die eigene Bestzeit.
create or replace function public.hindernisparkour2_submit(p_anon_id text, p_seconds double precision)
returns table (rank bigint, total bigint, best_seconds double precision, my_seconds double precision)
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.hindernisparkour2_runs as t (anon_id, seconds)
  values (p_anon_id, p_seconds)
  on conflict (anon_id) do update
    set seconds = least(t.seconds, excluded.seconds),
        updated_at = now();

  return query
  with me as (
    select r.seconds from public.hindernisparkour2_runs r where r.anon_id = p_anon_id
  )
  select
    (select count(*) + 1 from public.hindernisparkour2_runs r, me where r.seconds < me.seconds)::bigint,
    (select count(*) from public.hindernisparkour2_runs)::bigint,
    (select min(r.seconds) from public.hindernisparkour2_runs r)::double precision,
    (select me.seconds from me)::double precision;
end;
$$;

-- Bestzeit aller Teilnehmer (Startseite).
create or replace function public.hindernisparkour2_best_time()
returns double precision
language sql
stable
security definer
set search_path = public
as $$
  select min(seconds) from public.hindernisparkour2_runs;
$$;

-- Teilnehmerzahl (Startseite) – eine Zeile pro Gerät.
create or replace function public.hindernisparkour2_participant_count()
returns bigint
language sql
stable
security definer
set search_path = public
as $$
  select count(*) from public.hindernisparkour2_runs;
$$;

grant execute on function public.hindernisparkour2_submit(text, double precision) to anon, authenticated;
grant execute on function public.hindernisparkour2_best_time() to anon, authenticated;
grant execute on function public.hindernisparkour2_participant_count() to anon, authenticated;
