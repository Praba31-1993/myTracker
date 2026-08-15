# Life Goal Tracker — Full Stack + Excel Storage

## Architecture

- Frontend: HTML/CSS/JavaScript (the existing 6-level Life Goal Tracker UI)
- Backend: Node.js + Express
- Persistent storage: Microsoft Excel workbook (`data/life-goals.xlsx`)
- Database: **None**
- Authentication: not included in this version

## Excel sheets

The backend maintains all application data in `data/life-goals.xlsx`:

- `Levels` — level title, period and purpose
- `Goals` — main goals
- `Activities` — sub-goals, status and completion
- `Habits` — schedules, start/end time and frequency
- `AppState` — application state values such as completion flags

## Run locally

```bash
npm install
npm start
```

Open:

```text
http://localhost:3000
```

The Excel workbook is automatically created on first start.

## API

- `GET /api/health` — server/storage status
- `GET /api/state` — load the complete application state from Excel
- `PUT /api/state` — save the complete application state to Excel
- `GET /api/download/excel` — download the current workbook

## Important limitation

Excel is a file-based data store. This design is suitable for a personal/small-team tracker on a single server. It is not equivalent to a database for high-concurrency or multi-server deployment.
