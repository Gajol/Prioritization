# User Guide for Centre Leads

## What this file is

Once a quarter, Management sends you a copy of this workbook, already named for
your centre. It's how your centre reports:

1. What teams you have.
2. What priorities each team is working on, and how you rank them.
3. How your people's time is split across teams.

Fill it in and send it back. No Power Pivot, no macros, no add-ins required — just
Excel.

## Before you start

Open the **Instructions** tab. Your centre's name and code are already filled
in (grey, locked) at the bottom — Management stamps these in when they
generate your copy of the workbook, so there's nothing for you to enter
there. If they look wrong, tell Management before you start entering data.

## The sheets

- **Grey-header tabs** (Centres, Position, ProblemSet, InitiativeType,
  AssistanceType, RatingLookup, Resources, Priority, Tactical, Initiative,
  Assistance) are Management's reference data. They're locked — you can look, but
  you can't type into them. If something you need isn't there (a person, a
  position, a priority), tell Management; don't try to add it yourself.
- **Blue-header tabs** are yours: Step 1, Step 2, Step 3.

## Step 1 — Teams

List every team your centre has:

| Column | What to enter |
|---|---|
| Team Name | Free text, must be unique |
| Resources Dedicated | Yes / No / Temp |
| Priority Type | Tactical / Initiative / Assistance |

A team only works on **one** Priority Type. If a team genuinely spans two types,
list it twice under two different names.

The last two columns (Priority Allocation Total, Resource Effort Total) fill in
automatically as you complete Steps 2 and 3 — don't type into them.

## Step 2 — Priorities & Ranking

For each team, list the priorities it's working on:

| Column | What to enter |
|---|---|
| Team Name | Pick from the dropdown (must already exist in Step 1) |
| Priority Title | Pick from the dropdown — **only priorities matching the team's Priority Type are offered** |
| Type (auto) | Fills in automatically from the team — don't type into it |
| Rank | A positive whole number, unique within the team (1 = highest priority) |
| Resourced | Yes / No — is the team actually putting resources against this? |
| Allocation % of Team Effort | What share of the team's total effort goes to this priority |
| Value/Risk (auto) | Fills in automatically once you pick a Priority Title — Management's own Risk (Tactical) or Value (Initiative/Assistance) rating for it, colour-coded. Use it while you rank: a #1 you've ranked "Minimal" is worth a second look. |

A team's Allocation % across all its priorities can't exceed 100% — Excel will
refuse an entry that would push it over.

## Step 3 — Resource Allocation

For each person working on a team:

| Column | What to enter |
|---|---|
| Resource | Pick the person from the dropdown |
| Position Title (auto) | Fills in automatically — don't type into it |
| Team Name | Pick from the dropdown |
| Allocation % of Person's Time | What share of this person's time goes to this team |

A person's Allocation % across all their teams can't exceed 100%.

## Reading the check columns

Several columns are grey and calculate automatically — a quick way to sanity-check
your entries before sending the file back:

- **Green** — totals to exactly 100%. Good.
- **Amber** — under 100%. Not necessarily wrong (a team might not be fully
  allocated yet), but worth a second look.
- **Red** — over 100%. Something's wrong; fix it before sending the file back.

The colour on **Value/Risk (auto)** (Step 2) means something different — it's
Management's Risk/Value band for that priority (green/light-green = low risk or
high value, red = high risk or low value, depending on the priority's Type), not
a 100%-total check. Use it to sanity-check your ranking, not your allocation.

## Troubleshooting

- **A cell shows "unknown team" or "unknown resource"** — the Team Name or
  Resource you picked doesn't match anything in Step 1 or the Resources list.
  Usually means a typo got in before you switched to the dropdown, or a team was
  renamed after you'd already referenced it elsewhere.
- **Excel refuses my entry with a pop-up** — that's the point: it's telling you the
  value would create a duplicate team name, a duplicate rank within a team, or push
  an allocation over 100%. Read the message; it names the specific problem.
- **I pasted several rows at once and something looks off** — pasting can skip the
  live validation checks. Glance at the check columns (they still recalculate) and
  fix anything red before sending the file back.
- **I don't see the person/position/priority I need** — that comes from
  Management's reference data, which you can't edit here. Ask Management to add it
  for the next cycle.

## Sending it back

Save the file and return it the way Management asked (typically via SharePoint).
Don't rename the file unless asked to — the filename and the locked Centre
Name/Code on the Instructions tab are set together when Management generates
your copy, and consolidation across all six centres relies on them staying in
sync.
