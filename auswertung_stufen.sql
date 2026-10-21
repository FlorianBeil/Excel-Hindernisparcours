-- Excel.Flo – Auswertung Hindernisparcours Teil 1 → Teil 2
--
-- Einmalig im Supabase-Dashboard ausführen: SQL Editor → New query → einfügen → Run.
-- Kann gefahrlos mehrfach ausgeführt werden.
--
-- Die Seiten (app.js, teil2/app.js) schreiben seit 2026-10-21 in public.events:
--   portal 'hindernisparcours', exercise_id 'teil1' | 'teil2', session_id = Geräte-ID
--   (dieselbe wie fürs Ranking, in Teil 1 und 2 gleich)
--   challenge_start    Parcours gestartet (auch jeder Neustart)
--   challenge_finish   Parcours beendet            detail.sekunden
--   unlock_submit      Teil-2-Anmeldung abgeschickt (nur Teil 1)
-- Gezählt wird je Gerät: wer zehnmal spielt, zählt pro Woche einmal.

-- 1. 'hindernisparcours' als erlaubten Bereich ergänzen (die übrigen müssen drinbleiben)
alter table public.events drop constraint if exists events_portal_check;
alter table public.events add constraint events_portal_check
  check (portal in ('funktionen', 'pivot', 'powerquery', 'einstufungstest', 'kostenlos', 'light', 'hindernisparcours'));

-- 2. Auswertung
create or replace view public."Auswertung Stufen Hindernisparcours"
with (security_invoker = true) as
with schritte as (
  select
    date_trunc('week', created_at at time zone 'Europe/Berlin')::date as woche,
    session_id,
    case
      when exercise_id = 'teil1' and event = 'challenge_start'  then 1
      when exercise_id = 'teil1' and event = 'challenge_finish' then 2
      when exercise_id = 'teil1' and event = 'unlock_submit'    then 3
      when exercise_id = 'teil2' and event = 'challenge_start'  then 4
      when exercise_id = 'teil2' and event = 'challenge_finish' then 5
    end as schritt
  from public.events
  where portal = 'hindernisparcours'
),
zahlen as (
  select
    woche,
    count(distinct session_id) filter (where schritt = 1) as s1_start,
    count(distinct session_id) filter (where schritt = 2) as s1_ende,
    count(distinct session_id) filter (where schritt = 3) as anmeldung,
    count(distinct session_id) filter (where schritt = 4) as s2_start,
    count(distinct session_id) filter (where schritt = 5) as s2_ende
  from schritte
  where schritt is not null
  group by rollup (woche)
)
select
  coalesce(to_char(woche, 'DD.MM.YYYY'), 'Gesamt') as "Woche ab",
  s1_start  as "Stufe 1 gestartet",
  s1_ende   as "Stufe 1 beendet",
  anmeldung as "Anmeldung Stufe 2",
  s2_start  as "Stufe 2 gestartet",
  s2_ende   as "Stufe 2 beendet",
  round(100.0 * s1_ende / nullif(s1_start, 0), 1)   as "Stufe 1 beendet in %",
  round(100.0 * anmeldung / nullif(s1_ende, 0), 1)  as "Anmeldung in % der Beendeten",
  round(100.0 * s2_start / nullif(anmeldung, 0), 1) as "Stufe 2 gestartet in % der Anmeldungen",
  round(100.0 * s2_ende / nullif(s2_start, 0), 1)   as "Stufe 2 beendet in %"
from zahlen
order by woche is not null, woche desc;

revoke all on public."Auswertung Stufen Hindernisparcours" from anon, authenticated;
