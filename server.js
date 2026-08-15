const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, 'data');
const EXCEL_FILE = path.join(DATA_DIR, 'life-goals.xlsx');
const DEFAULT_FILE = path.join(DATA_DIR, 'default-state.json');

app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(ROOT, 'public')));

let writeQueue = Promise.resolve();

function clone(v) { return JSON.parse(JSON.stringify(v)); }

function defaultState() {
  const seed = JSON.parse(fs.readFileSync(DEFAULT_FILE, 'utf8'));
  return {
    levels: seed.levels,
    store: {
      habits: JSON.stringify([
        ['05:00','05:15','Wake up / planning','Daily'],
        ['05:15','06:15','Exercise / workout','Daily'],
        ['06:15','07:15','Yoga / flexibility','Daily'],
        ['07:15','08:00','Reading / learning','Daily'],
        ['08:00','09:00','Communication practice / recording','Daily'],
        ['20:00','21:00','Technical learning / business learning','Weekdays'],
        ['21:00','21:30','Reading / planning','Daily'],
        ['21:30','21:45','Daily review','Daily']
      ])
    }
  };
}

function stateToRows(state) {
  const levels = state.levels || [];
  const store = state.store || {};
  const levelRows = levels.map((l, i) => ({
    levelIndex: i + 1, title: l.title, period: l.period, purpose: l.purpose
  }));
  const goalRows = [];
  const activityRows = [];
  levels.forEach((l, li) => {
    (l.goals || []).forEach((g, gi) => {
      const goalId = `${li}-${gi}`;
      goalRows.push({ levelIndex: li + 1, goalIndex: gi + 1, goalId, title: g[0], description: g[1] });
      (g[2] || []).forEach((task, ti) => {
        activityRows.push({
          levelIndex: li + 1,
          goalIndex: gi + 1,
          goalId,
          activityIndex: ti + 1,
          activity: task,
          status: store[`S${li}G${gi}T${ti}`] || 'Not Started',
          completed: store[`L${li}G${gi}T${ti}`] === '1' ? 'Yes' : 'No'
        });
      });
    });
  });

  const habits = (() => {
    try { return JSON.parse(store.habits || '[]'); } catch { return []; }
  })();
  const habitRows = habits.map((h, i) => ({
    habitIndex: i + 1,
    completed: store[`HD${i}`] === '1' ? 'Yes' : 'No',
    startTime: h[0] || '', endTime: h[1] || '', activity: h[2] || '', frequency: h[3] || 'Daily'
  }));

  const stateRows = Object.entries(store).map(([key, value]) => ({ key, value: String(value ?? '') }));
  return { levelRows, goalRows, activityRows, habitRows, stateRows };
}

function rowsToState(workbook) {
  const fallback = defaultState();
  if (!workbook.Sheets.Levels || !workbook.Sheets.Goals) return fallback;
  const levelRows = XLSX.utils.sheet_to_json(workbook.Sheets.Levels, { defval: '' });
  const goalRows = XLSX.utils.sheet_to_json(workbook.Sheets.Goals, { defval: '' });
  const activityRows = XLSX.utils.sheet_to_json(workbook.Sheets.Activities, { defval: '' });
  const habitRows = workbook.Sheets.Habits ? XLSX.utils.sheet_to_json(workbook.Sheets.Habits, { defval: '' }) : [];
  const stateRows = workbook.Sheets.AppState ? XLSX.utils.sheet_to_json(workbook.Sheets.AppState, { defval: '' }) : [];

  const levels = levelRows.sort((a,b) => Number(a.levelIndex)-Number(b.levelIndex)).map(r => ({
    title: r.title, period: r.period, purpose: r.purpose, goals: []
  }));
  goalRows.forEach(r => {
    const li = Number(r.levelIndex) - 1;
    if (!levels[li]) return;
    levels[li].goals.push([r.title, r.description, []]);
  });
  levels.forEach(l => l.goals.forEach((g, gi) => {
    const li = levels.indexOf(l);
    const rows = activityRows
      .filter(r => Number(r.levelIndex) === li + 1 && Number(r.goalIndex) === gi + 1)
      .sort((a,b) => Number(a.activityIndex)-Number(b.activityIndex));
    rows.forEach((r, ti) => {
      g[2].push(r.activity);
    });
  }));

  const store = {};
  stateRows.forEach(r => { store[r.key] = String(r.value); });
  if (habitRows.length) {
    const habits = habitRows.sort((a,b)=>Number(a.habitIndex)-Number(b.habitIndex)).map((r,i)=>[
      r.startTime, r.endTime, r.activity, r.frequency
    ]);
    store.habits = JSON.stringify(habits);
    habitRows.forEach((r,i)=> { store[`HD${i}`] = String(r.completed).toLowerCase()==='yes' ? '1' : '0'; });
  }
  return { levels, store };
}

async function ensureWorkbook() {
  await fs.promises.mkdir(DATA_DIR, { recursive: true });
  try { await fs.promises.access(EXCEL_FILE); }
  catch { await writeExcel(defaultState()); }
}

async function readExcel() {
  await ensureWorkbook();
  const wb = XLSX.readFile(EXCEL_FILE);
  return rowsToState(wb);
}

function writeExcel(state) {
  writeQueue = writeQueue.then(async () => {
    await fs.promises.mkdir(DATA_DIR, { recursive: true });
    const wb = XLSX.utils.book_new();
    const rows = stateToRows(state);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.levelRows), 'Levels');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.goalRows), 'Goals');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.activityRows), 'Activities');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.habitRows), 'Habits');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.stateRows), 'AppState');
    XLSX.writeFile(wb, EXCEL_FILE);
  });
  return writeQueue;
}

app.get('/api/health', (req,res) => res.json({ ok: true, storage: 'Excel', database: false }));

app.get('/api/state', async (req,res,next) => {
  try { res.json(await readExcel()); } catch (e) { next(e); }
});

app.put('/api/state', async (req,res,next) => {
  try {
    if (!req.body || !Array.isArray(req.body.levels)) return res.status(400).json({ error: 'Invalid state payload.' });
    const state = { levels: req.body.levels, store: req.body.store || {} };
    await writeExcel(state);
    res.json({ ok: true, storage: 'Excel', file: 'data/life-goals.xlsx' });
  } catch (e) { next(e); }
});

app.get('/api/download/excel', async (req,res,next) => {
  try {
    await ensureWorkbook();
    res.download(EXCEL_FILE, 'life-goals.xlsx');
  } catch (e) { next(e); }
});

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: 'Server error', message: err.message });
});

app.get('/{*splat}', (req,res) => res.sendFile(path.join(ROOT, 'public', 'index.html')));

ensureWorkbook().then(() => {
  app.listen(PORT, () => console.log(`Life Goal Tracker running at http://localhost:${PORT}`));
}).catch(err => { console.error(err); process.exit(1); });
