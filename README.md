# Faculty Examination Timetable Planner

A web app for preparing faculty examination timetables from Excel course and room lists. The interface follows a three-session faculty timetable with courses and venues arranged under each date.

## Features

- Import an Excel workbook with **Courses** and **Rooms** sheets.
- Generate weekday and optional Saturday examination sessions.
- Prevent overlapping exams for courses sharing a named student group.
- Optionally allocate one or more available rooms to cover the stated student count, without using a room for overlapping exams.
- Produce a first draft with blank venue columns when room assignment is not ready.
- Review the day-by-day layout, print it, and export Excel sheets for the faculty layout, detailed assignments, and unscheduled courses.

The scheduling algorithm is greedy. It reports courses it cannot place; it does not prove that no alternative arrangement exists. Clash checking depends on accurate group names in the imported workbook. If students take courses across multiple groups, supply their actual registration information or define overlapping groups accordingly before relying on a timetable.

## Run locally

Requires Node.js 22.13 or newer. The project uses pnpm:

```bash
corepack enable
pnpm install
pnpm dev
```

Open the local URL printed by the development server. Use **Download Excel template** in the app for the expected workbook columns. The sample data is illustrative, not a faculty timetable.

## Workbook columns

| Sheet | Columns |
| --- | --- |
| Courses | Code, Title, Groups, Students, Duration (hours) |
| Rooms | Room, Capacity |

Separate multiple student groups in the **Groups** cell with semicolons. Use the exact same spelling wherever the same group takes another course. Course codes and room names must be unique. The Rooms sheet is part of the template even if you leave venue assignment off while drafting.

## Current scope

The planner runs in the browser and does not save imported data to a server. It schedules each listed course once within an available session. Faculty staff should verify registration lists, examination duration, venue availability, and the exported draft before publication.
