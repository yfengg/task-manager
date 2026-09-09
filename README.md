# Task & Deadline Manager

An Excel/VBA tool for tracking tasks and ranking them by priority. Tasks are entered 
through a form (urgency, difficulty, and importance scored 1–5) and logged to a sheet 
for review.

## How it works
- `frmDailyActivities` — the entry form used to add a new task
- `Module1.bas` — ranking logic that scores and sorts tasks by urgency, difficulty, 
  and importance

## Try it
Download `TaskManager.xlsm` and open it in Excel with macros enabled 
(File → Options → Trust Center → Enable Macros, or click "Enable Content" 
when Excel prompts you).

## Source
`src/` has the VBA exported as plain text (`.bas` / `.frm` files) so it's readable 
without opening Excel. The `.frx` file is a binary companion to the form — GitHub 
won't render it, that's expected.
