-- Excel.Flo – Auswertung Hindernisparcours mit Anmeldeseite (ab 2026-10)
--
-- Einmalig im Supabase-Dashboard ausführen: SQL Editor → New query → einfügen → Run.
-- Kann gefahrlos mehrfach ausgeführt werden. Die alte Ansicht
-- "Auswertung Stufen Hindernisparcours" (Teil 1 → Teil 2) bleibt für die alten Wochen stehen.
--
-- Ablauf seit dem Umbau: Anmeldeseite (/) → Bestätigungsmail → Parcours (/parcours/)
-- → Ergebnis mit Shortcut-Turbo. Ereignisse in public.events, portal 'hindernisparcours':
--   exercise_id 'anmeldung':  signup_view, signup_submit
--   exercise_id 'teil1':      challenge_start, challenge_finish, turbo_video, turbo_click
-- Gezählt wird je Gerät und Woche. Achtung: Wer sich am Handy anmeldet und am Computer
-- spielt, zählt als zwei Geräte – die Quoten zwischen Anmeldung und Parcours sind deshalb
-- nur ein grober Anhaltspunkt.

create or replace view public."Auswertung Hindernisparcours"
with (security_invoker = true) as
with schritte as (
  select
    date_trunc('week', created_at at time zone 'Europe/Berlin')::date as woche,
    session_id,
    case
      when exercise_id = 'anmeldung' and event = 'signup_view'      then 1
      when exercise_id = 'anmeldung' and event = 'signup_submit'    then 2
      when exercise_id = 'teil1'     and event = 'challenge_start'  then 3
      when exercise_id = 'teil1'     and event = 'challenge_finish' then 4
      when exercise_id = 'teil1'     and event = 'turbo_video'      then 5
      when exercise_id = 'teil1'     and event = 'turbo_click'      then 6
    end as schritt
  from public.events
  where portal = 'hindernisparcours'
),
zahlen as (
  select
    woche,
    count(distinct session_id) filter (where schritt = 1) as besuch,
    count(distinct session_id) filter (where schritt = 2) as anmeldung,
    count(distinct session_id) filter (where schritt = 3) as start,
    count(distinct session_id) filter (where schritt = 4) as ende,
    count(distinct session_id) filter (where schritt = 5) as video,
    count(distinct session_id) filter (where schritt = 6) as turbo
  from schritte
  where schritt is not null
  group by rollup (woche)
)
select
  coalesce(to_char(woche, 'DD.MM.YYYY'), 'Gesamt') as "Woche ab",
  besuch    as "Anmeldeseite besucht",
  anmeldung as "Angemeldet",
  round(100.0 * anmeldung / nullif(besuch, 0), 1) as "Anmeldung in % der Besucher",
  start     as "Parcours gestartet",
  ende      as "Parcours beendet",
  round(100.0 * ende / nullif(start, 0), 1)       as "Beendet in %",
  video     as "Video gestartet",
  turbo     as "Klick Shortcut-Turbo",
  round(100.0 * turbo / nullif(ende, 0), 1)       as "Klick in % der Beendeten"
from zahlen
order by woche is not null, woche desc;

revoke all on public."Auswertung Hindernisparcours" from anon, authenticated;
