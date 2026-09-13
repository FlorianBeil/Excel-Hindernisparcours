-- Teilnehmerzahl für die Startseite des Hindernisparcours.
-- Einmal im Supabase-Dashboard unter "SQL Editor" ausführen.
-- Die Tabelle hindernisparkour_runs bleibt für die öffentliche API gesperrt;
-- diese Funktion gibt nur die Anzahl der Zeilen (= Teilnehmer, eine Zeile pro Gerät) heraus.

create or replace function public.hindernisparkour_participant_count()
returns bigint
language sql
stable
security definer
set search_path = public
as $$
  select count(*) from public.hindernisparkour_runs;
$$;

grant execute on function public.hindernisparkour_participant_count() to anon, authenticated;
