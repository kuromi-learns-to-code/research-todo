/**
 * Research To-do — Google Apps Script 백엔드
 *
 * 데이터는 이 스크립트가 붙어 있는 Google 스프레드시트의 두 시트에 저장된다.
 *   Tasks   : 할일 (✅ / ➡️)
 *   Records : 분석 기록 (▷ 기록 / ✍🏻 기록에 딸린 추가 정보, 여러 개면 next 칸에 한 줄에 하나씩)
 *
 * 제공 기능
 *   1) 웹앱 페이지 (Index.html)                      → /exec?key=KEY
 *   2) 아이폰 단축어용 노트 텍스트                     → /exec?api=notes&key=KEY[&date=yyyy-MM-dd][&header=1]
 *   3) 전체 백업(JSON)                                → /exec?api=backup&key=KEY
 *   4) GitHub Pages 앱용 JSON API (POST)               → doPost
 */

const TZ = 'Asia/Seoul';
const H_TODO = '𝑹𝒆𝒔𝒆𝒂𝒓𝒄𝒉 𝑻𝒐-𝒅𝒐';
const H_REC = '𝑨𝒏𝒂𝒍𝒚𝒔𝒊𝒔 𝑹𝒆𝒄𝒐𝒓𝒅𝒔';
// 메모 기호 — 할일·단계: ✅ 완료 / ▶️ 진행중 / 💤 미룸(마감 지남), 기록 🔶 / 추가 정보 🔸
// (가져오기는 예전 기호 ✔︎ ➔ 😡 ➡️ ⏰ ▷ ✍🏻 ✓ · ☑ ☐ 도 인식)
const MARK = { done: '✅', open: '▶️', overdue: '💤', rec: '🔶', next: '🔸' };
const SCHEMA = {
  Tasks: ['id', 'project', 'title', 'status', 'created', 'done', 'updated', 'due', 'remind', 'repeat', 'until', 'nextId', 'steps', 'dueAuto', 'note', 'list'],
  Records: ['id', 'date', 'project', 'content', 'next', 'nextDone', 'deleted', 'updated'],
};
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
// 반복 규칙: d=매일, wd=평일, w:1,3=매주 월·수(0=일), m:15=매월 15일 (말일보다 크면 그 달 마지막 날)
const REPEAT_RE = /^(d|wd|w:[0-6](,[0-6])*|m:([1-9]|[12]\d|3[01]))$/;
// 딸림 할일(단계): steps 칸에 JSON [{"t":"내용","d":true/false}]. 메모에는 할일 아래 들여쓰기로 넘어감
const STEP_INDENT = '   ';   // 단계는 할일 아래 들여쓰기
// 예전 단계 기호
const isStepOn_ = function (c) { return c === '✓' || c === '☑' || c === '✔'; };
const isStepOff_ = function (c) { return c === '·' || c === '☐'; };
const REMIND_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/;   // 알림 시각 (한국시간, 예: 2026-11-09T14:30)
const WEEKDAY = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];

/* =====================================================================
 * 최초 1회: 편집기에서 setup 실행 → 실행 로그에 접속 키(KEY)가 찍힌다.
 * 다시 실행해도 기존 데이터는 지워지지 않는다.
 * ===================================================================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트의 [확장 프로그램 > Apps Script]에서 연 프로젝트에서 실행하세요.');
  Object.keys(SCHEMA).forEach(function (name) {
    const cols = SCHEMA[name];
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', ss.getId());
  let key = props.getProperty('KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('KEY', key);
  }
  Logger.log('접속 키(KEY): ' + key);
  return key;
}

/** 키가 유출됐다고 생각되면 실행. 새 키가 로그에 찍힌다 (페이지 북마크·단축어 URL도 바꿔야 함). */
function resetKey() {
  const key = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  PropertiesService.getScriptProperties().setProperty('KEY', key);
  Logger.log('새 접속 키(KEY): ' + key);
  return key;
}

/* =====================================================================
 * 웹 진입점
 * ===================================================================== */
/**
 * GitHub Pages 앱(index.html)이 호출하는 JSON API.
 * 요청 본문: {"fn": "apiLoad", "key": "...", "args": [...]}
 * 응답: {"ok": true, "data": ...} 또는 {"ok": false, "error": "..."}
 */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const args = [req.key].concat(Array.isArray(req.args) ? req.args : []);
    let data;
    switch (req.fn) {
      case 'apiLoad': data = apiLoad.apply(null, args); break;
      case 'apiSaveTask': data = apiSaveTask.apply(null, args); break;
      case 'apiSaveRecord': data = apiSaveRecord.apply(null, args); break;
      case 'apiNotes': data = apiNotes.apply(null, args); break;
      case 'apiImport': data = apiImport.apply(null, args); break;
      default: throw new Error('unknown fn: ' + req.fn);
    }
    out = { ok: true, data: data };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api) return apiGet_(p);
  const t = HtmlService.createTemplateFromFile('Index');
  t.key = String(p.key || '').replace(/[^A-Za-z0-9]/g, '');
  return t.evaluate()
    .setTitle('Research To-do')
    .setFaviconUrl('https://fonts.gstatic.com/s/e/notoemoji/latest/2705/512.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  // (Apps Script는 viewport 등 일부 meta 태그만 허용 → 앱 설정은 GitHub Pages 쪽 index.html에서 처리)
}

function apiGet_(p) {
  let body;
  try {
    auth_(p.key);
    const date = DATE_RE.test(p.date || '') ? p.date : today_();
    if (p.api === 'notes') {
      body = buildNotes_(date, p.header === '1');
    } else if (p.api === 'backup') {
      const json = JSON.stringify({
        exported: stamp_(),
        tasks: read_('Tasks').map(strip_),
        records: read_('Records').map(strip_),
      }, null, 1);
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    } else {
      body = 'ERROR: unknown api';
    }
  } catch (err) {
    body = 'ERROR: ' + err.message;
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.TEXT);
}

/* =====================================================================
 * 페이지에서 google.script.run 으로 호출하는 함수들 (첫 인자는 항상 key)
 * ===================================================================== */
function apiLoad(key) {
  auth_(key);
  return {
    today: today_(),
    url: ScriptApp.getService().getUrl(),
    tasks: read_('Tasks').filter(function (t) { return t.status !== 'deleted'; }).map(strip_),
    records: read_('Records').filter(function (r) { return r.deleted !== '1'; }).map(strip_),
  };
}

/** 할일 추가/수정. input: {id, project?, title?, status?, done?, due?, remind?, repeat?, until?, steps?, note?, list?} */
function apiSaveTask(key, input) {
  auth_(key);
  input = input || {};
  return withLock_(function () {
    const today = today_();
    const cur = input.id ? find_('Tasks', input.id) : null;
    const t = cur || { id: safeId_(input.id), project: '', title: '', status: 'open', created: today, done: '', due: '', remind: '', repeat: '', until: '', nextId: '', steps: '', dueAuto: '', note: '', list: '' };
    let spawned = null, removed = '';
    const prevProject = t.project;
    if (input.project != null) t.project = clean_(input.project);
    if (input.title != null) t.title = clean_(input.title);
    if (input.note != null) t.note = cleanMulti_(input.note);   // 할일 메모 (메모 앱으로는 넘기지 않음)
    // 목록: '' = Research To-do, 'backlog' = Backlog
    if (input.list != null) t.list = input.list === 'backlog' ? 'backlog' : '';
    // Backlog 항목에 태그를 새로 붙이거나 바꾸면 Research To-do로 이동 (같은 저장에서 목록을 직접 지정한 경우는 제외)
    let movedToTodo = false;
    if (cur && t.list === 'backlog' && input.list == null && t.project && t.project !== prevProject) {
      t.list = ''; movedToTodo = true;
    }
    if (input.due != null) { t.due = DATE_RE.test(input.due) ? input.due : ''; t.dueAuto = ''; }   // 직접 입력한 마감
    if (input.remind != null) t.remind = REMIND_RE.test(input.remind) ? input.remind : '';
    if (input.repeat != null) t.repeat = REPEAT_RE.test(input.repeat) ? input.repeat : '';
    if (input.until != null) t.until = DATE_RE.test(input.until) ? input.until : '';
    if (input.steps != null) {
      t.steps = stepsToStr_(input.steps);
      // 모든 단계를 체크하면 할일도 완료
      const st = stepsOf_(t.steps);
      if (st.length && st.every(function (x) { return x.d; }) && t.status === 'open' && !input.status) {
        input.status = 'done'; input.done = ''; t._autoDone = true;
      }
    }
    // 마감을 직접 넣지 않았으면 단계 마감 중 가장 늦은 날로 자동 설정 (단계 날짜가 바뀌면 따라감)
    if (!t.due || t.dueAuto === '1') {
      const sd = stepsOf_(t.steps).map(function (x) { return x.due; }).filter(Boolean).sort();
      if (sd.length) { t.due = sd[sd.length - 1]; t.dueAuto = '1'; }
      else if (t.dueAuto === '1') { t.due = ''; t.dueAuto = ''; }
    }
    if (t.repeat && !t.due) t.due = nextDue_(t.repeat, addDays_(today, -1), today) || today;   // 반복은 마감일 기준: 첫 회차 = 오늘 이후 첫 해당일
    if (t.repeat && t.due && t.status === 'open') alignToRule_(t);   // 마감일이 반복 요일/날짜와 어긋나면 맞춘다
    if (input.status && input.status !== t.status) {
      const wasOpen = t.status === 'open';
      if (['open', 'done', 'deleted'].indexOf(input.status) < 0) throw new Error('잘못된 상태값: ' + input.status);
      if (input.status === 'done') t.done = DATE_RE.test(input.done || '') ? input.done : today;
      if (input.status === 'open') t.done = '';
      t.status = input.status;
      // 반복 할일을 완료하면 다음 회차를 만든다
      if (wasOpen && t.status === 'done' && t.repeat && !t.nextId) {
        const nd = nextDue_(t.repeat, t.due || today, today);
        if (nd && (!t.until || nd <= t.until)) {
          spawned = {
            id: newId_(), project: t.project, title: t.title, status: 'open', created: today, done: '',
            due: nd, remind: shiftRemind_(t.remind, t.due, nd), repeat: t.repeat, until: t.until, nextId: '', updated: stamp_(),
            dueAuto: t.dueAuto, note: t.note, list: t.list,
            // 단계는 체크 해제해서 복사, 단계 마감도 회차 간격만큼 이동
            steps: stepsToStr_(stepsOf_(t.steps).map(function (x) {
              return { t: x.t, d: false, due: x.due && t.due ? shiftDate_(x.due, t.due, nd) : '' };
            })),
          };
          t.nextId = spawned.id;
        }
      }
      // 완료를 취소하면, 아직 손대지 않은 다음 회차는 지운다
      if (t.status === 'open' && t.nextId) {
        const nx = find_('Tasks', t.nextId);
        if (nx && nx.status === 'open') { nx.status = 'deleted'; nx.updated = stamp_(); write_('Tasks', nx); removed = nx.id; }
        t.nextId = '';
      }
    }
    if (!t.title) throw new Error('할 일 내용이 비어 있습니다.');
    t.updated = stamp_();
    write_('Tasks', t);
    if (spawned) appendRows_('Tasks', [spawned]);
    const autoDone = !!t._autoDone; delete t._autoDone;
    const out = strip_(t);
    if (movedToTodo) out._moved = true;
    if (autoDone) out._autoDone = true;
    if (spawned) out._next = strip_(spawned);
    if (removed) out._removed = removed;
    return out;
  });
}

/** 분석 기록 추가/수정. input: {id, date?, project?, content?, next?, nextDone?, deleted?} */
function apiSaveRecord(key, input) {
  auth_(key);
  input = input || {};
  return withLock_(function () {
    const cur = input.id ? find_('Records', input.id) : null;
    const r = cur || { id: safeId_(input.id), date: today_(), project: '', content: '', next: '', nextDone: '', deleted: '' };
    if (input.date != null && DATE_RE.test(input.date)) r.date = input.date;
    if (input.project != null) r.project = clean_(input.project);
    if (input.content != null) r.content = cleanMulti_(input.content);
    if (input.next != null) r.next = nextItems_(input.next).join('\n');
    if (input.nextDone != null) r.nextDone = String(input.nextDone);
    if (input.deleted != null) r.deleted = input.deleted ? '1' : '';
    if (!r.content && !r.next) throw new Error('기록 내용이 비어 있습니다.');
    r.updated = stamp_();
    write_('Records', r);
    return strip_(r);
  });
}

function apiNotes(key, date, withHeader) {
  auth_(key);
  return buildNotes_(DATE_RE.test(date || '') ? date : today_(), !!withHeader);
}

/** 기존 아이폰 메모 텍스트를 붙여넣어 데이터로 복원 */
function apiImport(key, text, defaultDate) {
  auth_(key);
  if (!DATE_RE.test(defaultDate || '')) defaultDate = today_();
  const days = parseNote_(String(text || ''), defaultDate);
  const res = { days: days.length, added: 0, completed: 0, records: 0, skipped: 0 };
  if (!days.length) return res;

  return withLock_(function () {
    const now = stamp_();
    const tasks = read_('Tasks').filter(function (t) { return t.status !== 'deleted'; });
    const recs = read_('Records').filter(function (r) { return r.deleted !== '1'; });
    const touched = [];
    const newTasks = [];
    const newRecs = [];
    const keyOf = function (p, t) { return (p + '|' + t).toLowerCase(); };
    const touch = function (t) { t.updated = now; if (t._row && touched.indexOf(t) < 0) touched.push(t); };

    days.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

    days.forEach(function (day) {
      day.tasks.forEach(function (it) {
        const k = keyOf(it.project, it.title);
        const same = tasks.filter(function (t) { return keyOf(t.project, t.title) === k; });
        const open = same.filter(function (t) { return t.status === 'open'; })[0];
        if (it.status === 'done') {
          if (open) {
            open.status = 'done';
            open.done = day.date;
            if (open.created > day.date) open.created = day.date;
            touch(open);
            res.completed++;
          } else if (same.some(function (t) { return t.done === day.date; })) {
            res.skipped++;
          } else {
            const n = { id: newId_(), project: it.project, title: it.title, status: 'done', created: day.date, done: day.date, updated: now, due: '', remind: '', steps: stepsToStr_(it.steps), note: it.note || '' };
            tasks.push(n); newTasks.push(n); res.added++;
          }
        } else {
          // ➡️ : 이미 열려 있거나, 이 날짜 이후에 완료된 같은 할일이 있으면 같은 건으로 본다
          const ongoing = open || same.filter(function (t) { return t.status === 'done' && t.done >= day.date; })[0];
          if (ongoing) {
            if (ongoing.created > day.date) { ongoing.created = day.date; touch(ongoing); }
            res.skipped++;
          } else {
            const n = { id: newId_(), project: it.project, title: it.title, status: 'open', created: day.date, done: '', updated: now, due: maxStepDue_(it.steps), dueAuto: maxStepDue_(it.steps) ? '1' : '', remind: '', steps: stepsToStr_(it.steps), note: it.note || '' };
            tasks.push(n); newTasks.push(n); res.added++;
          }
        }
      });
      day.records.forEach(function (it) {
        const k = [day.date, it.project, it.content, it.next].join('|');
        const dup = recs.some(function (r) { return [r.date, r.project, r.content, r.next].join('|') === k; });
        if (dup) { res.skipped++; return; }
        const n = {
          id: newId_(), date: day.date, project: it.project, content: it.content, next: it.next,
          nextDone: '', deleted: '', updated: now,
        };
        recs.push(n); newRecs.push(n); res.records++;
      });
    });

    touched.forEach(function (t) { write_('Tasks', t); });
    appendRows_('Tasks', newTasks);
    appendRows_('Records', newRecs);
    return res;
  });
}

/* =====================================================================
 * 노트 텍스트 생성
 *   ✅ : 그 날짜에 완료한 할일
 *   ➡️ : 그 날짜가 끝날 때 아직 열려 있던 할일 (다음 날로 이월)
 *   😡 : 그중 마감일이 이미 지난 할일 (밀린 일)
 * ===================================================================== */
function buildNotes_(date, withHeader) {
  const all = read_('Tasks');
  const spawnedIds = {};   // 반복 할일을 완료해서 자동으로 생긴 다음 회차
  all.forEach(function (t) { if (t.nextId) spawnedIds[t.nextId] = true; });
  const todo = all.filter(function (t) {
    if (t.status === 'deleted' || !t.created || t.created > date) return false;
    if (t.list === 'backlog') return false;   // Backlog는 메모에 넘기지 않음
    if (spawnedIds[t.id] && t.status === 'open' && t.due > date) return false;   // 아직 차례가 오지 않은 다음 회차
    if (t.status === 'open') return true;
    return t.status === 'done' && t.done >= date;
  }).map(function (t) {
    const mark = t.status === 'done' && t.done === date ? MARK.done
      : (t.due && t.due < date ? MARK.overdue : MARK.open);
    const lines = [mark + ' ' + label_(t.project, t.title)];
    // 할일 메모: 할일 바로 아래 들여쓴 🔸 줄 (여러 줄이면 줄마다)
    nextItems_(t.note).forEach(function (x) { lines.push(STEP_INDENT + MARK.next + ' ' + x); });
    stepsOf_(t.steps).forEach(function (x) {
      const sm = x.d ? MARK.done : (x.due && x.due < date ? MARK.overdue : MARK.open);
      lines.push(STEP_INDENT + sm + ' ' + x.t + stepDueLabel_(x.due));
    });
    return lines.join('\n');
  });

  const recs = [];
  read_('Records').forEach(function (r) {
    if (r.deleted === '1' || r.date !== date) return;
    if (r.content) recs.push(MARK.rec + ' ' + label_(r.project, r.content));
    nextItems_(r.next).forEach(function (x) { recs.push(STEP_INDENT + MARK.next + ' ' + x); });   // 🔸는 기록 아래 들여쓰기
  });

  const blocks = [];
  if (todo.length) blocks.push([H_TODO].concat(todo).join('\n'));
  if (recs.length) blocks.push([H_REC].concat(recs).join('\n'));
  const body = blocks.join('\n\n');
  // 날짜 제목 줄 바로 아래에 본문 (빈 줄 없음)
  return withHeader && body ? koDate_(date) + '\n' + body : body;
}

/* =====================================================================
 * 메모 텍스트 파서 (가져오기용)
 *   - "2026.11.06", "2026-11-06", "2026년 11월 6일" 같은 줄이 있으면 그 아래는 해당 날짜로 처리
 *   - ✅ / ➡️ / ▷ / ✍🏻 로 시작하는 줄을 인식, ▷·✍🏻 다음의 일반 줄은 이어지는 내용으로 붙인다
 * ===================================================================== */
function parseNote_(text, defaultDate) {
  const days = [];
  let date = defaultDate;
  let last = null;
  let lastField = '';
  let lastTask = null;

  const dayOf = function (d) {
    let x = days.filter(function (y) { return y.date === d; })[0];
    if (!x) { x = { date: d, tasks: [], records: [] }; days.push(x); }
    return x;
  };

  text.split(/\r?\n/).forEach(function (raw) {
    const line = raw.replace(/ /g, ' ').trim();
    if (!line) { lastField = ''; return; }
    const kind = markOf_(line);
    const n = line.normalize('NFKC');
    // 단계 줄: 들여쓴 ✅/▶️/💤 (예전: ✓ · ☑ ☐)
    const indented = /^[ \t\u00a0\u3000]/.test(raw);
    const oldStep = !kind && (isStepOn_(line.charAt(0)) || isStepOff_(line.charAt(0)));
    const isStep = lastTask && ((indented && (kind === 'done' || kind === 'open')) || oldStep);
    // 할일 아래 들여쓴 🔸 줄 → 할일 메모
    if (lastTask && indented && kind === 'next') {
      const nt = stripMark_(line);
      if (nt) lastTask.note = lastTask.note ? lastTask.note + '\n' + nt : nt;
      return;
    }
    if (isStep) {
      const st = (oldStep ? line.slice(1).replace(/^[\uFE0E\uFE0F]/, '') : stripMark_(line)).trim();
      const stepDone = oldStep ? isStepOn_(line.charAt(0)) : kind === 'done';
      if (st) {
        // "(~11/12)" 꼬리표 → 단계 마감 (연도는 메모 날짜 기준, 반년 넘게 앞서면 다음 해)
        const m = st.match(/^(.*?)\s*\(~(\d{1,2})\/(\d{1,2})\)$/);
        const step = { t: m ? m[1] : st, d: stepDone };
        if (m) {
          let y = Number(date.slice(0, 4));
          let due = y + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
          if (due < addDays_(date, -182)) due = (y + 1) + due.slice(4);
          step.due = due;
        }
        lastTask.steps.push(step);
      }
      return;
    }
    lastTask = null;

    if (!kind) {
      const dm = n.match(/^(\d{4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/);
      if (dm) { date = dm[1] + '-' + pad2_(dm[2]) + '-' + pad2_(dm[3]); last = null; lastField = ''; return; }
      if (/^research\s*to-?do/i.test(n) || /^analysis\s*records?/i.test(n)) { last = null; lastField = ''; return; }
      // ▷ 다음 줄은 기록에 줄바꿈으로, ✍🏻 다음 줄은 같은 ✍🏻 항목에 이어 붙인다
      if (last && lastField) last[lastField] += (lastField === 'next' ? ' ' : '\n') + line;
      return;
    }

    const day = dayOf(date);
    const body = stripMark_(line);
    if (kind === 'done' || kind === 'open') {
      const s = splitProj_(body);
      lastTask = { project: s.project, title: s.text, status: kind, steps: [], note: '' };
      day.tasks.push(lastTask);
      last = null; lastField = '';
    } else if (kind === 'rec') {
      const s = splitProj_(body);
      last = { project: s.project, content: s.text, next: '' };
      day.records.push(last);
      lastField = 'content';
    } else if (kind === 'next') {
      if (last && day.records.indexOf(last) >= 0) {
        last.next = last.next ? last.next + '\n' + body : body;
      } else {
        last = { project: '', content: '', next: body };
        day.records.push(last);
      }
      lastField = 'next';
    }
  });

  return days.filter(function (d) {
    d.tasks = d.tasks.filter(function (t) { return t.title; });
    return d.tasks.length || d.records.length;
  });
}

function markOf_(line) {
  const c = line.charAt(0);
  if (c === '✅' || c === '✔') return 'done';                                       // ✅ (예전 ✔︎)
  if (c === '▶' || line.indexOf('\uD83D\uDCA4') === 0 ||                          // ▶️ 💤
      c === '➔' || c === '➡' || c === '→' || c === '⏰' || line.indexOf('\uD83D\uDE21') === 0) return 'open';   // 예전 ➔ ➡️ ⏰ 😡
  if (line.indexOf('\uD83D\uDD36') === 0 || c === '▷' || c === '▸') return 'rec';     // 🔶 (예전 ▷)
  if (line.indexOf('\uD83D\uDD38') === 0 || c === '✍') return 'next';            // 🔸 (예전 ✍🏻)
  return '';
}

function stripMark_(line) {
  return line.replace(/^(?:[✅✔➔➡→⏰▷▶▸✍]|\uD83D[\uDE21\uDD36\uDD38\uDCA4])[\uFE0E\uFE0F]?(?:\uD83C[\uDFFB-\uDFFF])?[\uFE0E\uFE0F]?\s*/, '').trim();
}

function splitProj_(s) {
  const m = s.match(/^\[([^\]]+)\]\s*([\s\S]*)$/);
  return m ? { project: m[1].trim(), text: m[2].trim() } : { project: '', text: s.trim() };
}

/* =====================================================================
 * 시트 입출력
 * ===================================================================== */
let SS_CACHE_ = null;
function ss_() {
  if (SS_CACHE_) return SS_CACHE_;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  SS_CACHE_ = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return SS_CACHE_;
}

const HEADER_OK_ = {};
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(name + ' 시트가 없습니다. 편집기에서 setup()을 먼저 실행하세요.');
  if (!HEADER_OK_[name]) {
    // 코드 업데이트로 열이 추가된 경우 머리글을 자동으로 맞춘다 (기존 데이터는 그대로)
    const cols = SCHEMA[name];
    const head = sh.getRange(1, 1, 1, cols.length).getValues()[0];
    if (cols.some(function (c, i) { return String(head[i]) !== c; })) {
      sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    }
    HEADER_OK_[name] = true;
  }
  return sh;
}

function read_(name) {
  const sh = sheet_(name);
  const cols = SCHEMA[name];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().map(function (row, i) {
    const o = { _row: i + 2 };
    cols.forEach(function (c, j) {
      const v = row[j];
      o[c] = v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v == null ? '' : v);
    });
    return o;
  }).filter(function (o) { return o.id; });
}

function find_(name, id) {
  return read_(name).filter(function (o) { return o.id === id; })[0] || null;
}

function write_(name, obj) {
  const sh = sheet_(name);
  const cols = SCHEMA[name];
  const row = obj._row || sh.getLastRow() + 1;
  ensureRows_(sh, row);
  const rg = sh.getRange(row, 1, 1, cols.length);
  rg.setNumberFormat('@');
  rg.setValues([cols.map(function (c) { return obj[c] == null ? '' : String(obj[c]); })]);
  obj._row = row;
  return obj;
}

function appendRows_(name, objs) {
  if (!objs.length) return;
  const sh = sheet_(name);
  const cols = SCHEMA[name];
  const start = sh.getLastRow() + 1;
  ensureRows_(sh, start + objs.length - 1);
  const rg = sh.getRange(start, 1, objs.length, cols.length);
  rg.setNumberFormat('@');
  rg.setValues(objs.map(function (o) {
    return cols.map(function (c) { return o[c] == null ? '' : String(o[c]); });
  }));
}

function ensureRows_(sh, row) {
  const max = sh.getMaxRows();
  if (row > max) sh.insertRowsAfter(max, Math.max(100, row - max));
}

/* =====================================================================
 * 유틸
 * ===================================================================== */
function auth_(key) {
  const k = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!k || String(key || '') !== k) throw new Error('UNAUTHORIZED');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function stamp_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function safeId_(id) { const s = String(id || '').replace(/[^A-Za-z0-9_-]/g, ''); return s || newId_(); }
function clean_(s) { return String(s).replace(/\s+/g, ' ').trim(); }
function cleanMulti_(s) { return String(s).replace(/\r/g, '').replace(/[ \t]+$/gm, '').trim(); }
/** ✍🏻 항목들: next 칸에 한 줄에 하나씩 저장 */
function nextItems_(s) { return String(s || '').split('\n').map(function (x) { return x.trim(); }).filter(Boolean); }
function label_(project, text) { return project ? '[' + project + '] ' + text : text; }
function strip_(o) { const c = {}; Object.keys(o).forEach(function (k) { if (k !== '_row') c[k] = o[k]; }); return c; }
function pad2_(n) { return ('0' + n).slice(-2); }

/* ---------- 단계 ---------- */
function stepsOf_(v) {
  let a = v;
  if (typeof v === 'string') { try { a = v ? JSON.parse(v) : []; } catch (e) { a = []; } }
  if (!Array.isArray(a)) return [];
  return a.map(function (x) {
    const o = { t: clean_(x && x.t != null ? x.t : ''), d: !!(x && x.d) };
    if (x && DATE_RE.test(x.due || '')) o.due = x.due;
    return o;
  }).filter(function (x) { return x.t; }).slice(0, 50);
}
function maxStepDue_(steps) { const sd = stepsOf_(steps).map(function (x) { return x.due; }).filter(Boolean).sort(); return sd.length ? sd[sd.length - 1] : ''; }
function shiftDate_(d, fromRef, toRef) { return addDays_(d, Math.round((ymdToUTC_(toRef) - ymdToUTC_(fromRef)) / 864e5)); }
/** 메모용 단계 마감 표기: (~11/12) */
function stepDueLabel_(d) { return d ? ' (~' + Number(d.slice(5, 7)) + '/' + Number(d.slice(8)) + ')' : ''; }
function stepsToStr_(v) { const a = stepsOf_(v); return a.length ? JSON.stringify(a) : ''; }

/* ---------- 반복 ---------- */
function ymdToUTC_(s) { const p = s.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]); }
function utcToYmd_(t) { return new Date(t).toISOString().slice(0, 10); }
function addDays_(s, n) { return utcToYmd_(ymdToUTC_(s) + n * 864e5); }

/** 반복 규칙에 맞는 날 중, 이전 마감일보다 뒤이면서 오늘 이후인 첫 날 */
function nextDue_(rule, prevDue, today) {
  const after = prevDue >= today ? addDays_(prevDue, 1) : today;
  if (rule.indexOf('m:') === 0) {
    const want = Number(rule.slice(2));
    let p = after.split('-').map(Number);
    for (let k = 0; k < 24; k++) {
      const y = p[0] + Math.floor((p[1] - 1 + k) / 12), m = ((p[1] - 1 + k) % 12) + 1;
      const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
      const d = y + '-' + pad2_(m) + '-' + pad2_(Math.min(want, last));
      if (d >= after) return d;
    }
    return '';
  }
  const days = rule === 'd' ? [0, 1, 2, 3, 4, 5, 6] : rule === 'wd' ? [1, 2, 3, 4, 5]
    : rule.slice(2).split(',').map(Number);
  for (let i = 0; i < 14; i++) {
    const d = addDays_(after, i);
    if (days.indexOf(new Date(ymdToUTC_(d)).getUTCDay()) >= 0) return d;
  }
  return '';
}

/**
 * 마감일을 반복 규칙에 맞는 가장 가까운 날(마감일 포함 이후)로 옮긴다.
 * 알림: 원래 마감일보다 앞이면 같은 간격만큼 이동, 원래 마감일~새 마감일 사이면 새 마감일로, 그 뒤면 그대로.
 */
function alignToRule_(t) {
  const oldDue = t.due;
  const newDue = nextDue_(t.repeat, addDays_(oldDue, -1), oldDue);
  if (!newDue || newDue === oldDue) return;
  t.due = newDue;
  if (!t.remind) return;
  const p = t.remind.split('T');
  if (p[0] < oldDue) t.remind = shiftRemind_(t.remind, oldDue, newDue);
  else if (p[0] <= newDue) t.remind = newDue + 'T' + p[1];
}

/** 알림 시각을 마감일 이동만큼 옮긴다 (시각은 유지) */
function shiftRemind_(remind, fromDue, toDue) {
  if (!remind || !fromDue) return '';
  const p = remind.split('T');
  const gap = Math.round((ymdToUTC_(p[0]) - ymdToUTC_(fromDue)) / 864e5);
  return addDays_(toDue, gap) + 'T' + p[1];
}

function koDate_(ymd) {
  const p = ymd.split('-').map(Number);
  const wd = WEEKDAY[new Date(p[0], p[1] - 1, p[2]).getDay()];
  return p[0] + '. ' + pad2_(p[1]) + '. ' + pad2_(p[2]) + '. ' + wd;   // 예: 2026. 11. 06. 金曜日
}
