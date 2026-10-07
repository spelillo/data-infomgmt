/*
 * app.js — the course site: routing, pages, the quiz runner, submissions,
 * the gradebook and the quiz template builder.
 *
 * Everything runs in the browser. Progress and submissions are saved in this
 * browser's localStorage; students hand work in by exporting a receipt file,
 * which the teacher's gradebook re-grades from the work itself.
 */
(function (root) {
  'use strict';

  const { Formula: F, Workbook: W, Grader: G, Xlsx: X, Content: C, SheetView, QueryEditor } = root.DIM;
  const { h, dialog, toast, field, select, downloadText, readFile } = root.DIM.Dom;
  const app = document.getElementById('app');

  // ===========================================================================
  // Storage
  // ===========================================================================
  const KEY = 'dim-course-v1';
  const DEFAULT_STATE = () => ({
    student: null,
    work: {},          // in-progress work: { [workKey]: { wb, answers, mode, updated } }
    submissions: [],   // graded hand-ins
    attempts: {},      // running timed attempts: { [assessmentId]: { startedAt } }
    custom: [],        // teacher-built quizzes
    teacher: { receipts: [] },
    settings: { start: defaultStart(), policy: 'first' },
  });

  function defaultStart() {
    // the Monday on or after today (a teacher sets the real date in Settings)
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    while (d.getDay() !== 1) d.setDate(d.getDate() + 1);
    return isoDate(d);
  }

  let state = load();
  function load() {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return DEFAULT_STATE();
      const s = JSON.parse(raw);
      const d = DEFAULT_STATE();
      return { ...d, ...s, settings: { ...d.settings, ...(s.settings || {}) }, teacher: { ...d.teacher, ...(s.teacher || {}) } };
    } catch (e) {
      return DEFAULT_STATE();
    }
  }
  let saveTimer = null;
  function save(now) {
    clearTimeout(saveTimer);
    const write = () => {
      try { localStorage.setItem(KEY, JSON.stringify(state)); }
      catch (e) { toast('Could not save progress in this browser (storage is full or blocked). Download your workbook to keep it.', 'error'); }
    };
    if (now) write(); else saveTimer = setTimeout(write, 400);
  }
  addEventListener('beforeunload', () => save(true));

  // ===========================================================================
  // Assessment registry
  // ===========================================================================
  function registry() {
    const map = new Map();
    for (const l of C.lessons) {
      (l.practice || []).forEach((p) => map.set(l.id + '/' + p.id, { kind: 'practice', a: p, lesson: l, id: l.id + '/' + p.id }));
      if (l.assignment) map.set(l.assignment.id, { kind: l.assignment.category === 'project' ? 'project' : 'assignment', a: l.assignment, lesson: l, id: l.assignment.id });
    }
    C.quizzes.forEach((q) => map.set(q.id, { kind: 'quiz', a: q, lesson: lessonOf(q.week), id: q.id }));
    C.exams.forEach((e) => map.set(e.id, { kind: 'exam', a: e, lesson: lessonOf(e.week), id: e.id }));
    state.custom.forEach((q) => map.set(q.id, { kind: 'quiz', a: q, lesson: lessonOf(q.week), id: q.id, custom: true }));
    return map;
  }
  const lessonOf = (week) => C.lessons.find((l) => l.week === Number(week));
  const find = (id) => registry().get(id);

  // ===========================================================================
  // Schedule
  // ===========================================================================
  function isoDate(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
  function parseISO(s) { const [y, m, d] = String(s).split('-').map(Number); return new Date(y, (m || 1) - 1, d || 1); }
  function weekStart(week) { const d = parseISO(state.settings.start); d.setDate(d.getDate() + (week - 1) * 7); return d; }
  function dayOf(week, offset, hour = 0, min = 0) { const d = weekStart(week); d.setDate(d.getDate() + offset); d.setHours(hour, min, 0, 0); return d; }
  function dueDate(entry) {
    const week = entry.a.week || (entry.lesson && entry.lesson.week);
    if (!week) return null;
    if (entry.kind === 'assignment' || entry.kind === 'project') return dayOf(week, 6, 23, 59); // Sunday 11:59 pm
    if (entry.kind === 'quiz') return dayOf(week, 0, 8, 0);   // Monday
    if (entry.kind === 'exam') return dayOf(week, 4, 8, 0);   // Friday
    return null;
  }
  const fmtDate = (d) => d ? d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' }) : '';
  const fmtDue = (entry) => {
    const d = dueDate(entry);
    if (!d) return '';
    if (entry.kind === 'assignment' || entry.kind === 'project') return 'Due ' + fmtDate(d) + ', 11:59 pm';
    return (entry.kind === 'exam' ? 'Exam day ' : 'Given ') + fmtDate(d);
  };
  function currentWeek() {
    const diff = Math.floor((Date.now() - parseISO(state.settings.start).getTime()) / (7 * 864e5));
    return diff < 0 ? 0 : Math.min(15, diff + 1);
  }

  // ===========================================================================
  // Submissions and scoring
  // ===========================================================================
  const submissionsFor = (id) => state.submissions.filter((s) => s.assessmentId === id).sort((a, b) => a.submittedAt.localeCompare(b.submittedAt));
  function countedSubmission(id) {
    const subs = submissionsFor(id);
    if (!subs.length) return null;
    return state.settings.policy === 'best' ? subs.reduce((a, b) => (b.percent > a.percent ? b : a)) : subs[0];
  }
  function totalPoints(a) {
    const auto = (a.questions || []).reduce((s, q) => s + (q.points !== undefined ? q.points : 1), 0);
    return auto + (a.manualPoints || 0);
  }
  const pct = (x) => (Math.round(x * 10) / 10).toFixed(1).replace(/\.0$/, '') + '%';

  // Weighted course grade from counted submissions (rubric points are pending until scored).
  function courseGrade(subsFor) {
    const reg = registry();
    const groups = { assignment: [], quiz: [], exam1: [], exam2: [], exam3: [], project: [] };
    for (const [id, e] of reg) {
      if (e.custom || e.kind === 'practice') continue;
      const s = subsFor(id);
      if (!s) continue;
      const p = s.manualScore !== undefined && s.manualScore !== null ? (s.score + s.manualScore) / s.max * 100 : s.autoPercent ?? s.percent;
      if (e.kind === 'assignment') groups.assignment.push(p);
      else if (e.kind === 'project') groups.project.push(p);
      else if (e.kind === 'quiz') groups.quiz.push(p);
      else if (e.kind === 'exam') groups['exam' + e.a.unit].push(p);
    }
    if (groups.quiz.length > 1) groups.quiz.splice(groups.quiz.indexOf(Math.min(...groups.quiz)), 1); // lowest quiz dropped
    const weights = { assignment: 30, quiz: 15, exam1: 15, exam2: 15, exam3: 15, project: 10 };
    let wsum = 0, total = 0;
    const parts = {};
    for (const [k, list] of Object.entries(groups)) {
      if (!list.length) { parts[k] = null; continue; }
      const avg = list.reduce((a, b) => a + b, 0) / list.length;
      parts[k] = avg;
      total += avg * weights[k];
      wsum += weights[k];
    }
    return { overall: wsum ? total / wsum : null, parts, weights };
  }

  // ===========================================================================
  // Layout
  // ===========================================================================
  const NAV = [['#/', 'Home', 'home'], ['#/lessons', 'Lessons', 'lessons'], ['#/work', 'Assignments & Quizzes', 'work'], ['#/templates', 'Quiz Templates', 'templates'], ['#/grades', 'Grades', 'grades']];

  function layout(active, content, opts = {}) {
    const cw = currentWeek();
    const header = h('header', { class: 'site-header' },
      h('div', { class: 'topbar' },
        h('a', { class: 'brand', href: '#/' }, h('span', { class: 'brand-mark', 'aria-hidden': 'true' }, gridIcon()), h('span', {}, h('strong', {}, C.course.title), h('small', {}, C.course.grade))),
        h('button', { class: 'nav-toggle', type: 'button', 'aria-label': 'Menu', onClick: (e) => e.currentTarget.closest('.site-header').classList.toggle('open') }, '☰'),
        h('nav', { class: 'mainnav', 'aria-label': 'Main' }, NAV.map(([href, label, key]) => h('a', { href, class: active === key ? 'active' : null, 'aria-current': active === key ? 'page' : null }, label))),
        h('button', { class: 'student-chip', type: 'button', onClick: () => signIn(), title: 'Who is working on this computer' }, state.student ? '👤 ' + state.student.name : 'Sign in')),
      h('div', { class: 'banner' },
        h('div', { class: 'banner-inner' },
          h('div', { class: 'banner-text' }, h('p', { class: 'eyebrow' }, 'Lessons by week'), h('p', { class: 'banner-sub' }, cw ? 'This is week ' + cw + ' of 15.' : 'Class starts ' + fmtDate(weekStart(1)) + '.')),
          h('div', { class: 'weekstrip', role: 'list' }, C.course.units.map((u) => h('div', { class: 'unit-group', role: 'listitem' },
            h('span', { class: 'unit-label' }, 'Unit ' + u.id),
            h('div', { class: 'unit-weeks' }, u.weeks.map((w) => {
              const l = lessonOf(w);
              return h('a', { href: '#/week/' + w, class: 'wchip ' + u.color + (opts.week === w ? ' active' : '') + (cw === w ? ' now' : ''), title: 'Week ' + w + ': ' + (l ? l.title : ''), 'aria-label': 'Week ' + w + ': ' + (l ? l.title : '') }, String(w));
            })))))),
      ));
    const main = h('main', { id: 'main', class: 'container', tabindex: '-1' }, content);
    const footer = h('footer', { class: 'site-footer' }, h('div', { class: 'container' },
      h('p', {}, C.course.title + ' · ' + C.course.grade),
      h('p', { class: 'muted' }, 'Your work saves automatically in this browser. Hand work in by submitting it here; your teacher collects the receipt file. ', h('a', { href: '#/settings' }, 'Settings'), ' · ', h('a', { href: '#/help' }, 'How this site works'))));
    app.replaceChildren(h('a', { class: 'skip', href: '#main', onClick: (e) => { e.preventDefault(); main.focus(); } }, 'Skip to content'), header, main, footer);
    // on narrow screens the week strip scrolls: bring the current week into view
    const strip = header.querySelector('.weekstrip');
    const chip = strip && (strip.querySelector('.wchip.active') || strip.querySelector('.wchip.now'));
    if (chip && strip.scrollWidth > strip.clientWidth) strip.scrollLeft = chip.offsetLeft - strip.offsetLeft - strip.clientWidth / 2 + chip.offsetWidth / 2;
  }

  function gridIcon() {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('width', '28'); s.setAttribute('height', '28');
    s.innerHTML = '<rect x="2" y="2" width="20" height="20" rx="4" fill="currentColor"/><path d="M2 9h20M2 15h20M9 2v20M15 2v20" stroke="#fff" stroke-width="1.6"/>';
    return s;
  }

  const crumb = (...parts) => h('nav', { class: 'crumbs', 'aria-label': 'Breadcrumb' }, parts.map((p, i) => [i ? h('span', { 'aria-hidden': 'true' }, ' / ') : null, Array.isArray(p) ? h('a', { href: p[0] }, p[1]) : h('span', {}, p)]));
  const badge = (text, kind = '') => h('span', { class: 'badge ' + kind }, text);
  const kindLabel = { assignment: 'Assignment', project: 'Project', quiz: 'Quiz', exam: 'Exam', practice: 'Practice' };

  // ===========================================================================
  // Student sign-in and settings
  // ===========================================================================
  function signIn(then) {
    const name = h('input', { type: 'text', value: state.student ? state.student.name : '', placeholder: 'First and last name', autofocus: true, maxlength: 60 });
    const period = h('input', { type: 'text', value: state.student ? state.student.period || '' : '', placeholder: 'e.g. 3', maxlength: 12 });
    dialog('Who is working?', h('div', {},
      h('p', { class: 'muted' }, 'Your name goes on every submission and receipt. It is stored only in this browser.'),
      field('Name', name), field('Class period (optional)', period)), [
      { label: 'Cancel' },
      { label: 'Save', primary: true, onClick: () => {
        const n = name.value.trim();
        if (!n) { toast('Type your name first.', 'warn'); return false; }
        state.student = { name: n, period: period.value.trim() };
        save(true);
        // update the header in place: re-rendering would throw away work in progress
        document.querySelectorAll('.student-chip').forEach((b) => { b.textContent = '👤 ' + n; });
        if (!then && /^#?\/?(grades|settings)?$/.test(location.hash)) route();
        if (then) then();
      } },
    ]);
  }
  const requireStudent = (fn) => (state.student ? fn() : signIn(fn));

  function pageSettings() {
    const start = h('input', { type: 'date', value: state.settings.start });
    const policy = select([['first', 'First attempt counts'], ['best', 'Best attempt counts']], state.settings.policy);
    layout('', h('div', { class: 'narrow' },
      crumb(['#/', 'Home'], 'Settings'),
      h('h1', {}, 'Settings'),
      h('div', { class: 'card' },
        h('h2', {}, 'Calendar'),
        field('First Monday of the semester', start, 'Due dates, quiz days and exam days are counted from this date.'),
        field('When a quiz or assignment is submitted more than once', policy),
        h('button', { class: 'btn btn-primary', type: 'button', onClick: () => { state.settings.start = start.value || state.settings.start; state.settings.policy = policy.value; save(true); toast('Settings saved.'); route(); } }, 'Save settings')),
      h('div', { class: 'card' },
        h('h2', {}, 'This browser\'s data'),
        h('p', {}, 'Progress, submissions and custom quizzes live in this browser only. Back them up or move them to another computer with a backup file.'),
        h('div', { class: 'btn-row' },
          h('button', { class: 'btn', type: 'button', onClick: () => downloadText(JSON.stringify(state), 'course-backup-' + isoDate(new Date()) + '.json', 'application/json') }, 'Download backup'),
          fileButton('Restore backup', '.json', async (file) => {
            try { const s = JSON.parse(await readFile(file)); if (!s || !Array.isArray(s.submissions)) throw new Error('not a backup'); state = { ...DEFAULT_STATE(), ...s }; save(true); toast('Backup restored.'); route(); }
            catch (e) { toast('That file is not a course backup.', 'error'); }
          }),
          h('button', { class: 'btn btn-danger', type: 'button', onClick: () => dialog('Erase everything?', h('p', {}, 'This deletes all saved work, submissions and custom quizzes in this browser. Download a backup first if you might need them.'), [{ label: 'Cancel' }, { label: 'Erase', primary: true, onClick: () => { state = DEFAULT_STATE(); save(true); route(); } }]) }, 'Erase all data'))),
    ));
  }

  function fileButton(label, accept, onFile, opts = {}) {
    const input = h('input', { type: 'file', accept, class: 'sr-only', multiple: !!opts.multiple, onChange: async (e) => {
      const files = [...e.target.files];
      e.target.value = '';
      if (opts.multiple) await onFile(files); else if (files[0]) await onFile(files[0]);
    } });
    return h('label', { class: 'btn' + (opts.primary ? ' btn-primary' : '') }, input, label);
  }

  // ===========================================================================
  // Pages: Home
  // ===========================================================================
  function pageHome() {
    const cw = currentWeek();
    const reg = registry();
    const lesson = lessonOf(cw || 1);
    const upcoming = [...reg.values()].filter((e) => e.kind !== 'practice' && !e.custom && dueDate(e) && dueDate(e) >= new Date(Date.now() - 864e5) && !countedSubmission(e.id))
      .sort((a, b) => dueDate(a) - dueDate(b)).slice(0, 4);
    const done = state.submissions.filter((s) => !s.practice).length;
    layout('home', [
      h('section', { class: 'hero' },
        h('div', {},
          h('p', { class: 'eyebrow' }, 'Course 1 · ' + C.course.grade),
          h('h1', {}, C.course.title),
          h('p', { class: 'lead' }, C.course.description),
          h('div', { class: 'btn-row' },
            h('a', { class: 'btn btn-primary btn-lg', href: '#/week/' + (cw || 1) }, cw ? 'Go to week ' + cw : 'Start with week 1'),
            h('a', { class: 'btn btn-lg', href: '#/work' }, 'Assignments & Quizzes'))),
        h('div', { class: 'hero-card card' },
          h('h2', {}, cw ? 'This week: week ' + cw : 'Before class starts'),
          lesson ? h('p', {}, h('a', { href: '#/week/' + lesson.week }, lesson.title)) : null,
          lesson ? h('p', { class: 'muted' }, lesson.summary) : null,
          upcoming.length ? h('ul', { class: 'due-list' }, upcoming.map((e) => h('li', {}, badge(kindLabel[e.kind], e.kind), ' ', h('a', { href: hrefFor(e) }, e.a.title), h('span', { class: 'muted' }, ' · ' + fmtDue(e))))) : h('p', { class: 'muted' }, 'Nothing due right now.'),
          h('p', { class: 'muted small' }, done + ' item' + (done === 1 ? '' : 's') + ' submitted from this browser.'))),
      h('section', {},
        h('h2', {}, 'Three units, fifteen weeks'),
        h('div', { class: 'grid3' }, C.course.units.map((u) => h('a', { class: 'card unit-card ' + u.color, href: '#/week/' + u.weeks[0] },
          h('p', { class: 'eyebrow' }, 'Unit ' + u.id + ' · Weeks ' + u.weeks[0] + '–' + u.weeks[u.weeks.length - 1]),
          h('h3', {}, u.title),
          h('ol', { start: u.weeks[0], class: 'unit-list' }, u.weeks.map((w) => h('li', {}, lessonOf(w).title))),
          h('p', { class: 'muted small' }, 'Exam ' + u.id + ' on Friday of week ' + u.weeks[u.weeks.length - 1] + ', covering these weeks only.'))))),
      h('section', { class: 'grid2' },
        h('div', { class: 'card' }, h('h2', {}, 'How each week works'), h('dl', { class: 'rhythm' }, C.course.rhythm.map(([d, t]) => [h('dt', {}, d), h('dd', {}, t)]))),
        h('div', { class: 'card' }, h('h2', {}, 'How you are graded'),
          h('table', { class: 'simple' }, h('thead', {}, h('tr', {}, h('th', {}, 'Category'), h('th', { class: 'num' }, 'Weight'))),
            h('tbody', {}, C.course.grading.map(([c, , w]) => h('tr', {}, h('td', {}, c), h('td', { class: 'num' }, w + '%'))))),
          h('p', { class: 'muted small' }, 'Workbooks are graded automatically the moment you submit, with feedback on every task.'))),
      h('section', { class: 'card' }, h('h2', {}, 'By the end of the course you will be able to'), h('ul', { class: 'outcomes' }, C.course.outcomes.map((o) => h('li', {}, o)))),
      h('section', { class: 'card' }, h('h2', {}, 'Tools'), h('table', { class: 'simple' },
        h('thead', {}, h('tr', {}, h('th', {}, 'Tool'), h('th', {}, 'When'), h('th', {}, 'Notes'))),
        h('tbody', {}, C.course.tools.map((t) => h('tr', {}, t.map((x) => h('td', {}, x))))))),
    ]);
  }

  function hrefFor(e) {
    if (e.kind === 'practice') return '#/week/' + e.lesson.week + '/' + e.a.id;
    if (e.kind === 'assignment' || e.kind === 'project') return '#/assignment/' + e.id;
    return '#/' + (e.kind === 'exam' ? 'exam' : 'quiz') + '/' + e.id;
  }

  // ===========================================================================
  // Pages: Lessons overview
  // ===========================================================================
  function pageLessons() {
    layout('lessons', [
      h('h1', {}, 'Lessons'),
      h('p', { class: 'lead' }, 'Every week has detailed notes, interactive Excel workbooks you can check as you go, a Sunday assignment, and (most weeks) a Monday quiz on the week before.'),
      C.course.units.map((u) => h('section', { class: 'unit-section' },
        h('h2', { class: 'unit-head ' + u.color }, 'Unit ' + u.id + ': ' + u.title),
        h('div', { class: 'week-cards' }, u.weeks.map((w) => {
          const l = lessonOf(w);
          const prog = practiceProgress(l);
          return h('a', { class: 'card week-card', href: '#/week/' + w },
            h('div', { class: 'week-num ' + u.color }, h('small', {}, 'Week'), h('span', {}, String(w))),
            h('div', {},
              h('h3', {}, l.title),
              h('p', { class: 'muted' }, l.summary),
              h('div', { class: 'tags' },
                badge((l.practice || []).length + ' workbook' + ((l.practice || []).length === 1 ? '' : 's')),
                l.assignment ? badge(l.assignment.title, 'assignment') : null,
                l.quiz ? badge('Mon: ' + find(l.quiz).a.title.split(' · ')[0], 'quiz') : null,
                l.exam ? badge('Fri: Exam ' + u.id, 'exam') : null,
                prog.total ? badge(prog.done + '/' + prog.total + ' checked', prog.done === prog.total ? 'ok' : '') : null)));
        })))),
    ]);
  }

  function practiceProgress(l) {
    const total = (l.practice || []).length;
    const done = (l.practice || []).filter((p) => { const w = state.work['practice:' + l.id + '/' + p.id]; return w && w.lastPercent === 100; }).length;
    return { total, done };
  }

  // ===========================================================================
  // Pages: Week
  // ===========================================================================
  function pageWeek(week, focus) {
    const l = lessonOf(week);
    if (!l) return notFound();
    const unit = C.course.units.find((u) => u.id === l.unit);
    const sections = [
      ['notes', 'Notes'], ['terms', 'Key terms'], ...(l.practice || []).map((p, i) => ['practice-' + p.id, 'Workbook ' + (i + 1)]),
      ...(l.assignment ? [['assignment', 'Assignment']] : []), ...(l.quiz || l.exam ? [['assess', 'Quiz & exam']] : []),
    ];
    const prev = lessonOf(week - 1), next = lessonOf(week + 1);
    layout('lessons', [
      crumb(['#/lessons', 'Lessons'], 'Unit ' + l.unit, 'Week ' + week),
      h('header', { class: 'lesson-head ' + unit.color },
        h('p', { class: 'eyebrow' }, 'Week ' + week + ' · Unit ' + l.unit + ': ' + unit.title),
        h('h1', {}, l.title),
        h('p', { class: 'lead' }, l.summary)),
      h('div', { class: 'lesson-layout' },
        h('aside', { class: 'lesson-toc', 'aria-label': 'On this page' },
          h('p', { class: 'toc-title' }, 'On this page'),
          h('ul', {}, sections.map(([id, label]) => h('li', {}, h('a', { href: '#/week/' + week + '/' + id, onClick: (e) => { e.preventDefault(); scrollToId(id); history.replaceState(null, '', '#/week/' + week + '/' + id); } }, label)))),
          dueBox(l)),
        h('article', { class: 'lesson-body' },
          h('section', { class: 'card objectives' }, h('h2', {}, 'By the end of this week you can'), h('ul', { class: 'checks' }, l.objectives.map((o) => h('li', {}, o)))),
          h('section', { id: 'notes', class: 'notes' }, l.sections.map((s) => h('div', { class: 'note-section' }, h('h2', {}, s.heading), h('div', { class: 'prose', html: s.html })))),
          h('section', { id: 'terms', class: 'card' }, h('h2', {}, 'Key terms'), h('dl', { class: 'terms' }, l.keyTerms.map(([t, d]) => [h('dt', {}, t), h('dd', {}, d)]))),
          (l.practice || []).map((p) => h('section', { id: 'practice-' + p.id, class: 'practice' }, practiceBlock(l, p))),
          l.assignment ? h('section', { id: 'assignment', class: 'card assignment-card' },
            h('p', { class: 'eyebrow' }, kindLabel[l.assignment.category === 'project' ? 'project' : 'assignment'] + ' · ' + totalPoints(l.assignment) + ' points · ' + fmtDue(find(l.assignment.id))),
            h('h2', {}, l.assignment.title),
            h('div', { class: 'prose', html: l.assignment.html }),
            statusLine(l.assignment.id),
            h('a', { class: 'btn btn-primary', href: '#/assignment/' + l.assignment.id }, 'Open the assignment')) : null,
          (l.quiz || l.exam) ? h('section', { id: 'assess', class: 'card' }, h('h2', {}, 'Quiz and exam'),
            l.quiz ? assessRow(find(l.quiz), 'Monday of this week. Covers week ' + find(l.quiz).a.covers + '.') : h('p', { class: 'muted' }, week === 1 ? 'No quiz in week 1.' : 'No quiz this Monday: it is the Monday after an exam.'),
            l.exam ? assessRow(find(l.exam), 'Friday of this week. Covers weeks ' + find(l.exam).a.covers[0] + '–' + find(l.exam).a.covers.slice(-1)[0] + ' only.') : null) : null,
          h('nav', { class: 'pager' },
            prev ? h('a', { class: 'btn', href: '#/week/' + prev.week }, '← Week ' + prev.week + ': ' + prev.title) : h('span'),
            next ? h('a', { class: 'btn', href: '#/week/' + next.week }, 'Week ' + next.week + ': ' + next.title + ' →') : h('span')))),
    ], { week });
    if (focus) setTimeout(() => scrollToId(focus), 30);
  }

  function scrollToId(id) {
    const el = document.getElementById(id);
    if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); if (id.startsWith('practice-')) { const b = el.querySelector('.open-wb'); if (b) b.focus(); } }
  }

  function dueBox(l) {
    const items = [];
    if (l.quiz) items.push(['Mon', find(l.quiz)]);
    if (l.exam) items.push(['Fri', find(l.exam)]);
    if (l.assignment) items.push(['Sun', find(l.assignment.id)]);
    if (!items.length) return null;
    return h('div', { class: 'due-box' }, h('p', { class: 'toc-title' }, 'This week'), h('ul', {}, items.map(([d, e]) => h('li', {}, h('span', { class: 'day' }, d), h('a', { href: hrefFor(e) }, e.a.title.replace(/^Quiz \d+ · |^Exam \d+ · /, (m) => m.replace(' · ', ': '))), statusMini(e.id)))));
  }

  function statusMini(id) {
    const s = countedSubmission(id);
    return s ? h('span', { class: 'mini-score' }, pct(s.percent)) : null;
  }
  function statusLine(id) {
    const s = countedSubmission(id);
    if (!s) return h('p', { class: 'muted' }, 'Not submitted yet.');
    return h('p', { class: 'status-ok' }, 'Submitted ' + new Date(s.submittedAt).toLocaleString() + ' · auto-graded ' + s.score + '/' + s.autoMax + ' (' + pct(s.autoPercent) + ')' + (s.manualMax ? ' · ' + s.manualMax + ' rubric points scored by your teacher' : ''));
  }
  function assessRow(e, note) {
    return h('div', { class: 'assess-row' },
      h('div', {}, badge(kindLabel[e.kind], e.kind), ' ', h('strong', {}, e.a.title), h('p', { class: 'muted' }, note + ' ' + e.a.minutes + ' minutes · ' + totalPoints(e.a) + ' points.'), statusLine(e.id)),
      h('a', { class: 'btn btn-primary', href: hrefFor(e) }, countedSubmission(e.id) ? 'Review' : 'Open'));
  }

  // ---------------------------------------------------------------------------
  // Practice workbook block (on the week page)
  // ---------------------------------------------------------------------------
  function practiceBlock(l, p) {
    const id = l.id + '/' + p.id;
    const key = 'practice:' + id;
    const wrap = h('div', { class: 'card practice-card' });
    const work = state.work[key];
    const head = h('div', { class: 'practice-head' },
      h('div', {}, h('p', { class: 'eyebrow' }, 'Practice · not graded · ' + totalPoints(p) + ' checks'), h('h2', {}, p.title), h('p', {}, p.intro)),
      work && work.lastPercent !== undefined ? h('div', { class: 'score-pill ' + (work.lastPercent === 100 ? 'ok' : '') }, pct(work.lastPercent)) : null);
    const body = h('div', { class: 'practice-body' });
    const open = () => {
      body.replaceChildren(assessmentWorkspace({ a: p, key, kind: 'practice', fileName: p.title.replace(/^Workbook /, ''), onGraded: (report) => {
        state.work[key] = { ...(state.work[key] || {}), lastPercent: report.percent };
        save();
      } }));
      openBtn.remove();
    };
    const openBtn = h('button', { class: 'btn btn-primary open-wb', type: 'button', onClick: open }, p.workbook ? 'Open workbook' : 'Start');
    wrap.append(head, openBtn, body);
    if (work) setTimeout(open, 0); // reopen work already started
    return wrap;
  }

  // ===========================================================================
  // The workspace: tasks + questions + workbook + check/submit
  // ===========================================================================
  /*
   * opts: { a, key, kind, fileName, onGraded(report, ctx), submitLabel,
   *         submit: (ctx) => void (graded hand-in) , locked }
   */
  function assessmentWorkspace(opts) {
    const a = opts.a;
    const key = opts.key;
    const saved = state.work[key] || {};
    const answers = { ...(saved.answers || {}) };
    const hasWorkbook = !!a.workbook;
    const wbQuestions = (a.questions || []).filter(G.isWorkbookQuestion);
    const pageQuestions = (a.questions || []).filter((q) => !G.isWorkbookQuestion(q));
    let mode = saved.mode === 'upload' ? 'upload' : 'web';
    let uploaded = null; // { wb, name }
    let view = null;
    let wb = null;

    const fresh = () => { const w = G.buildTemplate(a); w.sources = { ...(w.sources || {}) }; return w; };
    function loadWb() {
      if (!hasWorkbook) return null;
      if (saved.wb) { try { const w = W.Workbook.fromJSON(saved.wb); return w; } catch (e) { /* fall through */ } }
      return fresh();
    }
    function persist() {
      state.work[key] = { ...(state.work[key] || {}), answers, mode, wb: wb ? wb.toJSON() : undefined, updated: new Date().toISOString() };
      save();
    }

    const root_ = h('div', { class: 'workspace' });
    const itemEls = new Map();

    // Task list (workbook questions)
    const taskList = wbQuestions.length ? h('ol', { class: 'tasks' }, wbQuestions.map((q) => {
      const li = h('li', { class: 'task', 'data-q': q.id },
        h('div', { class: 'task-main' },
          h('span', { class: 'task-prompt' }, q.prompt || q.id),
          h('span', { class: 'task-meta' }, (q.sheet ? q.sheet + (q.cell || q.range ? '!' + (q.cell || q.range) : '') + ' · ' : '') + plural(q.points ?? 1, 'pt'))),
        h('div', { class: 'feedback', 'aria-live': 'polite' }));
      li.addEventListener('click', (e) => { if (e.target.closest('button,a')) return; if (view && q.sheet && wb.getSheet(q.sheet) && (q.cell || q.range)) { view.switchSheet(q.sheet); view.select(q.cell || String(q.range).split(':')[0]); } });
      itemEls.set(q.id, li);
      return li;
    })) : null;

    // Page questions
    const qList = pageQuestions.length ? h('ol', { class: 'questions' }, pageQuestions.map((q, i) => {
      const li = h('li', { class: 'question', 'data-q': q.id }, h('span', { class: 'q-num', 'aria-hidden': 'true' }, String(wbQuestions.length + i + 1)), questionInput(q, answers, () => { persist(); }), h('div', { class: 'feedback', 'aria-live': 'polite' }));
      itemEls.set(q.id, li);
      return li;
    })) : null;

    // Workbook area with web / upload modes
    const wbHost = h('div', { class: 'wb-host' });
    const uploadHost = h('div', { class: 'upload-host' });
    const modeTabs = hasWorkbook ? h('div', { class: 'mode-tabs', role: 'tablist' },
      modeTab('web', 'Work in the browser'), modeTab('upload', 'Upload from Excel (.xlsx)'),
      h('button', { class: 'btn btn-sm', type: 'button', onClick: () => X.download(X.write(fresh(), { title: a.title }), safeFile(a.title) + '.xlsx'), title: 'A blank copy of the starting workbook to complete in Excel' }, '⬇ Blank template (.xlsx)'),
      sourcesNote()) : null;
    function sourcesNote() {
      const srcs = Object.keys((a.workbook && a.workbook.sources) || {});
      if (!srcs.length) return null;
      return h('button', { class: 'btn btn-sm', type: 'button', onClick: () => srcs.forEach((f) => downloadText(a.workbook.sources[f], f, 'text/csv')), title: 'The data files this task imports: ' + srcs.join(', ') }, '⬇ Data files (' + srcs.length + ')');
    }
    function modeTab(m, label) {
      return h('button', { type: 'button', role: 'tab', class: 'mode-tab' + (mode === m ? ' active' : ''), 'aria-selected': mode === m ? 'true' : 'false', onClick: () => { mode = m; persist(); renderMode(); } }, label);
    }
    function renderMode() {
      if (!hasWorkbook) return;
      modeTabs.querySelectorAll('.mode-tab').forEach((b, i) => { const on = (i === 0 ? 'web' : 'upload') === mode; b.classList.toggle('active', on); b.setAttribute('aria-selected', on ? 'true' : 'false'); });
      wbHost.hidden = mode !== 'web';
      uploadHost.hidden = mode !== 'upload';
      if (mode === 'web' && !view) {
        wb = loadWb();
        view = SheetView.mount(wbHost, wb, {
          tools: a.tools || 'basic',
          fileName: opts.fileName || a.title,
          highlights: wbQuestions.filter((q) => q.cell || q.range).map((q) => ({ sheet: q.sheet, range: q.cell || q.range, label: q.prompt })),
          onChange: (w) => { wb = w; persist(); },
          onReset: () => { wb = fresh(); persist(); return wb; },
        });
      }
      if (mode === 'upload') renderUpload();
    }
    function renderUpload() {
      uploadHost.replaceChildren(
        h('div', { class: 'upload-box' },
          h('p', {}, h('strong', {}, 'Using Excel on a computer? '), 'Download the blank template, complete it in Excel, save, and upload the .xlsx here. It is graded exactly like the web workbook.'),
          h('ul', { class: 'small muted' },
            h('li', {}, 'Keep the sheet names the same. Name any new sheets exactly as the task says.'),
            h('li', {}, 'Tables, data validation, conditional formatting, charts, PivotTables and Power Query queries in your file are all read and checked.')),
          h('div', { class: 'btn-row' }, fileButton(uploaded ? 'Replace file' : 'Choose .xlsx file', '.xlsx', async (file) => {
            try {
              const buf = await readFile(file, 'buffer');
              const w = await X.read(buf);
              w.sources = { ...((a.workbook && a.workbook.sources) || {}), ...(w.sources || {}) };
              uploaded = { wb: w, name: file.name };
              toast('Read ' + file.name + ': ' + w.sheets.length + ' sheet(s).');
              renderUpload();
            } catch (e) {
              toast('Could not read that file: ' + e.message, 'error');
            }
          }, { primary: true })),
          uploaded ? h('p', { class: 'status-ok' }, '✓ ' + uploaded.name + ' · sheets: ' + uploaded.wb.sheets.map((s) => s.name).join(', ')) : h('p', { class: 'muted' }, 'No file chosen yet.')));
    }

    // Actions
    const reportHost = h('div', { class: 'report-host', 'aria-live': 'polite' });
    const actions = h('div', { class: 'ws-actions' });
    const currentWb = () => (mode === 'upload' ? (uploaded ? uploaded.wb : null) : wb);

    function runGrade(reveal) {
      if (hasWorkbook && mode === 'upload' && !uploaded) { toast('Choose your .xlsx file first, or switch to the web workbook.', 'warn'); return null; }
      if (hasWorkbook && mode === 'web' && view) view.commit();
      const target = currentWb();
      const report = G.grade(a, { workbook: target ? target.clone() : undefined, answers }, { reveal });
      showFeedback(report);
      return report;
    }
    function showFeedback(report) {
      report.items.forEach((it) => {
        const el = itemEls.get(it.id);
        if (!el) return;
        el.classList.remove('correct', 'partial', 'incorrect');
        el.classList.add(it.status);
        const fb = el.querySelector('.feedback');
        fb.replaceChildren(h('span', { class: 'fb-icon', 'aria-hidden': 'true' }, it.status === 'correct' ? '✓' : it.status === 'partial' ? '◐' : '✗'), h('span', {}, it.earned + '/' + it.points + ' · ' + it.feedback));
      });
    }

    if (opts.kind === 'practice') {
      actions.append(
        h('button', { class: 'btn btn-primary', type: 'button', onClick: () => {
          const r = runGrade(false);
          if (!r) return;
          reportHost.replaceChildren(scoreBanner(r, 'Check'));
          if (opts.onGraded) opts.onGraded(r);
          if (r.percent === 100) toast('Everything checks out. Nice work!', 'ok');
        } }, 'Check my work'),
        h('button', { class: 'btn', type: 'button', onClick: () => { const r = runGrade(true); if (r) reportHost.replaceChildren(scoreBanner(r, 'Check'), h('p', { class: 'muted small' }, 'Expected values are now shown in the feedback for anything not yet correct.')); } }, 'Show expected answers'));
    } else if (opts.submit) {
      actions.append(h('button', { class: 'btn btn-primary btn-lg', type: 'button', 'data-testid': 'submit', onClick: () => requireStudent(() => {
        if (hasWorkbook && mode === 'upload' && !uploaded) { toast('Choose your .xlsx file first, or switch to the web workbook.', 'warn'); return; }
        const unanswered = pageQuestions.filter((q) => isUnanswered(q, answers[q.id])).length;
        dialog('Submit ' + (kindLabel[opts.kind] || '').toLowerCase() + '?', h('div', {},
          h('p', {}, 'Your work will be graded now and saved as a submission for ' + state.student.name + '.'),
          unanswered ? h('p', { class: 'warn' }, unanswered + ' question' + (unanswered === 1 ? ' is' : 's are') + ' unanswered.') : null), [
          { label: 'Keep working' },
          { label: 'Submit', primary: true, onClick: () => doSubmit() },
        ]);
      }) }, opts.submitLabel || 'Submit'));
    }

    function doSubmit(auto) {
      if (hasWorkbook && mode === 'web' && view) view.commit();
      const target = currentWb();
      const report = G.grade(a, { workbook: target ? target.clone() : undefined, answers }, { reveal: true });
      showFeedback(report);
      opts.submit({ report, answers: { ...answers }, workbook: target ? target.toJSON() : null, mode, uploadName: uploaded ? uploaded.name : null, auto });
    }

    if (hasWorkbook) {
      root_.append(...[
        taskList ? h('div', { class: 'ws-tasks' }, h('h3', {}, 'Tasks in the workbook'), h('p', { class: 'muted small' }, 'Green-shaded cells are where your answers go. Click a task to jump to its cell.'), taskList) : null,
        modeTabs, wbHost, uploadHost,
        qList ? h('div', { class: 'ws-questions' }, h('h3', {}, 'Questions'), qList) : null,
        actions, reportHost].filter(Boolean));
      renderMode();
    } else {
      root_.append(...[qList ? h('div', { class: 'ws-questions' }, qList) : null, actions, reportHost].filter(Boolean));
    }
    root_.doSubmit = doSubmit;
    root_.destroy = () => { if (view) view.destroy(); };
    return root_;
  }

  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');
  const safeFile = (s) => String(s).replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '_') || 'workbook';
  function isUnanswered(q, v) {
    if (q.type === 'match') return !v || Object.keys(v).length < q.pairs.length;
    if (q.type === 'choice' && q.multi) return !v || !v.length;
    return v === undefined || v === null || v === '';
  }

  // One page question's input.
  function questionInput(q, answers, onChange) {
    const set = (v) => { answers[q.id] = v; onChange(); };
    const pts = h('span', { class: 'q-points' }, plural(q.points ?? 1, 'pt'));
    if (q.type === 'choice') {
      const name = 'q-' + q.id + '-' + Math.random().toString(36).slice(2, 7);
      const cur = answers[q.id];
      return h('fieldset', { class: 'q' }, h('legend', {}, q.prompt, ' ', pts, q.multi ? h('span', { class: 'muted small' }, ' (select all that apply)') : null),
        q.options.map((o, i) => h('label', { class: 'opt' },
          h('input', { type: q.multi ? 'checkbox' : 'radio', name, value: i, checked: q.multi ? [].concat(cur || []).map(Number).includes(i) : cur !== undefined && Number(cur) === i,
            onChange: (e) => {
              if (q.multi) {
                const s = new Set([].concat(answers[q.id] || []).map(Number));
                if (e.target.checked) s.add(i); else s.delete(i);
                set([...s].sort());
              } else set(i);
            } }), ' ', o)));
    }
    if (q.type === 'number' || q.type === 'text') {
      const id = 'q-' + q.id + '-' + Math.random().toString(36).slice(2, 7);
      return h('div', { class: 'q' }, h('label', { for: id }, q.prompt, ' ', pts),
        h('input', { id, type: 'text', inputmode: q.type === 'number' ? 'decimal' : null, value: answers[q.id] ?? '', autocomplete: 'off', onInput: (e) => set(e.target.value), class: q.type === 'number' ? 'num-input' : 'text-input' }));
    }
    if (q.type === 'match') {
      const cur = answers[q.id] || {};
      // definitions in a stable shuffled order
      const order = q.pairs.map((p, i) => i).sort((x, y) => hashStr(q.id + x) - hashStr(q.id + y));
      return h('fieldset', { class: 'q' }, h('legend', {}, q.prompt, ' ', pts),
        h('div', { class: 'match' }, q.pairs.map(([term], i) => h('label', { class: 'match-row' }, h('span', { class: 'match-term' }, term),
          select([['', 'Choose…'], ...order.map((j) => [j, q.pairs[j][1]])], cur[i] ?? '', { onChange: (e) => { const v = { ...(answers[q.id] || {}) }; if (e.target.value === '') delete v[i]; else v[i] = Number(e.target.value); set(v); } })))));
    }
    return h('p', {}, q.prompt);
  }
  function hashStr(s) { let x = 2166136261; for (let i = 0; i < s.length; i++) { x ^= s.charCodeAt(i); x = Math.imul(x, 16777619); } return x >>> 0; }

  function scoreBanner(r, label) {
    const cls = r.percent === 100 ? 'ok' : r.percent >= 70 ? 'mid' : 'low';
    return h('div', { class: 'score-banner ' + cls, role: 'status' },
      h('span', { class: 'score-big' }, pct(r.percent)),
      h('span', {}, label + ': ' + r.score + ' of ' + r.max + ' points. ' + r.items.filter((i) => i.status === 'correct').length + ' of ' + r.items.length + ' items fully correct.'));
  }

  // ===========================================================================
  // Pages: Assignment
  // ===========================================================================
  function pageAssignment(id) {
    const e = find(id);
    if (!e || (e.kind !== 'assignment' && e.kind !== 'project')) return notFound();
    const a = e.a;
    const subs = submissionsFor(id);
    const key = 'hw:' + id;
    const uploadState = { file: null };
    let ws = null;
    const container = h('div');
    const results = h('div');

    const renderResults = () => results.replaceChildren(subs.length ? submissionHistory(e, submissionsFor(id)) : '');

    const uploadField = a.upload ? h('div', { class: 'card' },
      h('h2', {}, 'Upload your file'),
      h('p', { class: 'muted' }, a.upload.label + '. Files up to 3 MB are packed into your receipt so your teacher can open them.'),
      fileButton('Choose file', a.upload.accept, (file) => {
        uploadState.file = file;
        fileLabel.textContent = '✓ ' + file.name + ' (' + Math.round(file.size / 1024) + ' KB)';
      }),
      h('p', { class: 'status-ok' }, '')) : null;
    const fileLabel = uploadField ? uploadField.querySelector('.status-ok') : null;

    ws = assessmentWorkspace({
      a, key, kind: e.kind, fileName: a.title, submitLabel: 'Submit assignment',
      submit: async (ctx) => {
        let file = null;
        if (uploadState.file) {
          file = { name: uploadState.file.name, size: uploadState.file.size, type: uploadState.file.type };
          if (uploadState.file.size <= 3 * 1024 * 1024) file.data = await toDataURL(uploadState.file);
        }
        const sub = recordSubmission(e, ctx, { file });
        showReceipt(e, sub);
        renderResults();
      },
    });

    layout('work', [
      crumb(['#/work', 'Assignments & Quizzes'], a.title),
      h('header', { class: 'page-head' },
        h('p', { class: 'eyebrow' }, kindLabel[e.kind] + ' · Week ' + e.lesson.week + ' · ' + fmtDue(e)),
        h('h1', {}, a.title),
        h('p', { class: 'muted' }, totalPoints(a) + ' points' + (a.manualPoints ? ' (' + (totalPoints(a) - a.manualPoints) + ' auto-graded, ' + a.manualPoints + ' scored by your teacher with the rubric)' : ', all auto-graded'))),
      h('div', { class: 'card prose', html: a.html }),
      a.rubric ? rubricCard(a) : null,
      uploadField,
      h('div', { class: 'card' }, h('h2', {}, a.workbook ? 'Your workbook' : 'Questions'), ws),
      results,
    ], { week: e.lesson.week });
    renderResults();
  }

  function rubricCard(a) {
    return h('details', { class: 'card rubric' }, h('summary', {}, 'Rubric'),
      h('table', { class: 'simple' }, h('thead', {}, h('tr', {}, h('th', {}, 'Criterion'), h('th', { class: 'num' }, 'Points'))),
        h('tbody', {}, a.rubric.map(([c, p]) => h('tr', {}, h('td', {}, c), h('td', { class: 'num' }, typeof p === 'number' ? p : p))))));
  }

  function toDataURL(file) {
    return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(file); });
  }

  function recordSubmission(e, ctx, extra = {}) {
    const r = ctx.report;
    const manualMax = e.a.manualPoints || 0;
    const sub = {
      id: 'sub-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      assessmentId: e.id, kind: e.kind, title: e.a.title,
      student: { ...state.student },
      submittedAt: new Date().toISOString(),
      score: r.score, autoMax: r.max, autoPercent: r.percent,
      max: r.max + manualMax, manualMax, manualScore: null,
      percent: manualMax ? Math.round((r.score / (r.max + manualMax)) * 1000) / 10 : r.percent,
      items: r.items.map(({ id, type, points, earned, status, feedback }) => ({ id, type, points, earned, status, feedback })),
      answers: ctx.answers, workbook: ctx.workbook, mode: ctx.mode, uploadName: ctx.uploadName,
      auto: !!ctx.auto, file: extra.file || null,
    };
    state.submissions.push(sub);
    save(true);
    return sub;
  }

  function receiptOf(sub) {
    return { type: 'dim-receipt', version: 1, course: C.course.title, ...sub };
  }
  function downloadReceipt(sub) {
    downloadText(JSON.stringify(receiptOf(sub)), safeFile(sub.student.name) + '_' + sub.assessmentId + '_' + sub.submittedAt.slice(0, 10) + '.dim.json', 'application/json');
  }

  function showReceipt(e, sub) {
    dialog('Submitted: ' + pct(sub.autoPercent) + ' auto-graded', h('div', {},
      h('p', {}, h('strong', {}, sub.title), ' · ' + sub.student.name),
      h('p', {}, 'Auto-graded score: ', h('strong', {}, sub.score + ' / ' + sub.autoMax), sub.manualMax ? ' · plus ' + sub.manualMax + ' rubric points your teacher will score' : ''),
      h('p', { class: 'muted' }, 'Your feedback is shown next to every task. To hand this in, download the receipt and turn it in where your teacher asks (it contains your work, so your teacher\'s gradebook re-checks it).')), [
      { label: 'Close' },
      { label: 'Download receipt', primary: true, onClick: () => { downloadReceipt(sub); } },
    ]);
  }

  function submissionHistory(e, subs) {
    const counted = countedSubmission(e.id);
    return h('div', { class: 'card' }, h('h2', {}, 'Your submissions'),
      h('table', { class: 'simple' },
        h('thead', {}, h('tr', {}, h('th', {}, '#'), h('th', {}, 'Submitted'), h('th', {}, 'Student'), h('th', { class: 'num' }, 'Auto score'), h('th', {}, ''), h('th', {}, ''))),
        h('tbody', {}, subs.map((s, i) => h('tr', { class: s === counted ? 'counted' : null },
          h('td', {}, String(i + 1)), h('td', {}, new Date(s.submittedAt).toLocaleString()), h('td', {}, s.student.name),
          h('td', { class: 'num' }, s.score + '/' + s.autoMax + ' (' + pct(s.autoPercent) + ')'),
          h('td', {}, s === counted ? badge('counts', 'ok') : ''),
          h('td', {}, h('button', { class: 'btn btn-sm', type: 'button', onClick: () => showReport(s) }, 'Feedback'), ' ', h('button', { class: 'btn btn-sm', type: 'button', onClick: () => downloadReceipt(s) }, 'Receipt')))))),
      h('p', { class: 'muted small' }, state.settings.policy === 'best' ? 'Your best attempt counts.' : 'Your first attempt counts; later attempts are for practice.'));
  }

  function showReport(s) {
    const e = find(s.assessmentId);
    const qs = new Map(((e && e.a.questions) || []).map((q) => [q.id, q]));
    dialog(s.title + ' · ' + pct(s.autoPercent), h('ol', { class: 'report-list' }, s.items.map((it) => h('li', { class: it.status },
      h('div', {}, h('strong', {}, (qs.get(it.id) && qs.get(it.id).prompt) || it.id)),
      h('div', { class: 'small' }, (it.status === 'correct' ? '✓ ' : it.status === 'partial' ? '◐ ' : '✗ ') + it.earned + '/' + it.points + ' · ' + it.feedback)))), [{ label: 'Close' }], { wide: true });
  }

  // ===========================================================================
  // Pages: Quiz / Exam runner
  // ===========================================================================
  let timerHandle = null;
  function pageQuiz(id) {
    const e = find(id);
    if (!e || (e.kind !== 'quiz' && e.kind !== 'exam')) return notFound();
    const a = e.a;
    const attempt = state.attempts[id];
    const subs = submissionsFor(id);
    const crumbs = crumb(['#/work', 'Assignments & Quizzes'], a.title);
    const head = h('header', { class: 'page-head' },
      h('p', { class: 'eyebrow' }, kindLabel[e.kind] + ' · ' + fmtDue(e) + (a.covers ? ' · covers ' + (Array.isArray(a.covers) ? 'weeks ' + a.covers[0] + '–' + a.covers.slice(-1)[0] : 'week ' + a.covers) : '')),
      h('h1', {}, a.title));

    if (!attempt) {
      layout('work', [crumbs, head,
        h('div', { class: 'card quiz-start' },
          h('p', { class: 'lead' }, a.intro || ''),
          h('ul', { class: 'facts' },
            h('li', {}, h('strong', {}, a.minutes + ' minutes'), ' once you start. The timer keeps running if you leave the page, and your work is submitted automatically when time runs out.'),
            h('li', {}, h('strong', {}, plural((a.questions || []).length, 'item')), ' worth ', h('strong', {}, totalPoints(a) + ' points'), '.'),
            a.workbook ? h('li', {}, 'Part of it happens in a workbook. Work in the browser, or download the blank template now and upload your finished .xlsx from Excel.') : null,
            h('li', {}, 'Graded automatically the moment you submit, with feedback on every item.')),
          h('div', { class: 'btn-row' },
            h('button', { class: 'btn btn-primary btn-lg', type: 'button', 'data-testid': 'start', onClick: () => requireStudent(() => {
              state.attempts[id] = { startedAt: Date.now() };
              delete state.work['quiz:' + id];
              save(true);
              route();
            }) }, subs.length ? 'Start a new attempt' : 'Start ' + kindLabel[e.kind].toLowerCase()),
            a.workbook ? h('button', { class: 'btn', type: 'button', onClick: () => X.download(X.write(G.buildTemplate(a), { title: a.title }), safeFile(a.title) + '.xlsx') }, '⬇ Blank template (.xlsx)') : null),
          subs.length && state.settings.policy === 'first' ? h('p', { class: 'muted small' }, 'You already submitted this. New attempts are recorded, but your first attempt counts.') : null),
        subs.length ? submissionHistory(e, subs) : null,
      ], { week: e.lesson && e.lesson.week });
      return;
    }

    const deadline = attempt.startedAt + a.minutes * 60000;
    const timer = h('div', { class: 'timer', role: 'timer', 'aria-live': 'off' });
    let ws;
    const submitNow = (auto) => {
      ws.doSubmit(auto);
    };
    ws = assessmentWorkspace({
      a, key: 'quiz:' + id, kind: e.kind, fileName: a.title, submitLabel: 'Submit ' + kindLabel[e.kind].toLowerCase(),
      submit: (ctx) => {
        clearInterval(timerHandle);
        const sub = recordSubmission(e, ctx);
        delete state.attempts[id];
        delete state.work['quiz:' + id];
        save(true);
        layout('work', [crumbs, head,
          scoreBanner({ percent: sub.autoPercent, score: sub.score, max: sub.autoMax, items: sub.items }, ctx.auto ? 'Time is up. Auto-submitted' : 'Submitted'),
          h('div', { class: 'btn-row' }, h('button', { class: 'btn btn-primary', type: 'button', onClick: () => downloadReceipt(sub) }, 'Download receipt'), h('a', { class: 'btn', href: '#/work' }, 'Back to Assignments & Quizzes')),
          h('div', { class: 'card' }, h('h2', {}, 'Feedback'), reportList(e, sub))]);
        toast('Submitted. Score: ' + pct(sub.autoPercent), 'ok');
      },
    });
    layout('work', [crumbs, head, h('div', { class: 'quiz-bar' }, timer, h('span', { class: 'muted' }, state.student ? state.student.name : '')), h('div', { class: 'card' }, ws)], { week: e.lesson && e.lesson.week });
    const tick = () => {
      const left = deadline - Date.now();
      if (left <= 0) { clearInterval(timerHandle); timer.textContent = '0:00'; toast('Time is up. Submitting your work…', 'warn'); submitNow(true); return; }
      const m = Math.floor(left / 60000), s = Math.floor((left % 60000) / 1000);
      timer.textContent = '⏱ ' + m + ':' + String(s).padStart(2, '0') + ' left';
      timer.classList.toggle('low', left < 120000);
    };
    clearInterval(timerHandle);
    timerHandle = setInterval(tick, 1000);
    tick();
  }

  function reportList(e, sub) {
    const qs = new Map((e.a.questions || []).map((q) => [q.id, q]));
    return h('ol', { class: 'report-list' }, sub.items.map((it) => h('li', { class: it.status },
      h('div', {}, h('strong', {}, (qs.get(it.id) && qs.get(it.id).prompt) || it.id)),
      h('div', { class: 'small' }, (it.status === 'correct' ? '✓ ' : it.status === 'partial' ? '◐ ' : '✗ ') + it.earned + '/' + it.points + ' · ' + it.feedback))));
  }

  // ===========================================================================
  // Pages: Assignments & Quizzes hub
  // ===========================================================================
  function pageWork() {
    const reg = [...registry().values()];
    const rows = (kinds) => reg.filter((e) => kinds.includes(e.kind) && !e.custom).sort((x, y) => (x.a.week || x.lesson.week) - (y.a.week || y.lesson.week));
    const table = (list) => h('table', { class: 'simple work-table' },
      h('thead', {}, h('tr', {}, h('th', {}, 'Week'), h('th', {}, 'Title'), h('th', {}, 'When'), h('th', { class: 'num' }, 'Points'), h('th', {}, 'Status'), h('th', {}, ''))),
      h('tbody', {}, list.map((e) => {
        const s = countedSubmission(e.id);
        const late = !s && dueDate(e) && dueDate(e) < new Date();
        return h('tr', {},
          h('td', {}, String(e.a.week || e.lesson.week)),
          h('td', {}, h('a', { href: hrefFor(e) }, e.a.title)),
          h('td', {}, fmtDue(e)),
          h('td', { class: 'num' }, String(totalPoints(e.a))),
          h('td', {}, s ? badge(pct(s.autoPercent) + (s.manualMax ? ' + rubric' : ''), s.autoPercent >= 70 ? 'ok' : 'warn') : late ? badge('Missing', 'late') : badge('Not started')),
          h('td', {}, h('a', { class: 'btn btn-sm', href: hrefFor(e) }, s ? 'Review' : e.kind === 'quiz' || e.kind === 'exam' ? 'Take' : 'Open')));
      })));
    const custom = reg.filter((e) => e.custom);
    layout('work', [
      h('h1', {}, 'Assignments & Quizzes'),
      h('p', { class: 'lead' }, 'Turn in homework, take quizzes and exams, and see your feedback. Everything with a workbook is graded automatically as soon as you submit.'),
      h('div', { class: 'grid3 howto' },
        h('div', { class: 'card' }, h('h3', {}, '1. Do the work'), h('p', {}, 'Work in the browser workbook (it saves as you go), or download the template and use Excel.')),
        h('div', { class: 'card' }, h('h3', {}, '2. Submit'), h('p', {}, 'Click Submit. Your work is graded instantly with feedback on every task.')),
        h('div', { class: 'card' }, h('h3', {}, '3. Hand in the receipt'), h('p', {}, 'Download the receipt and turn it in where your teacher asks. The teacher\'s gradebook re-grades it from your work.'))),
      h('section', { class: 'card' }, h('h2', { id: 'assignments' }, 'Weekly assignments (due Sunday 11:59 pm)'), table(rows(['assignment', 'project']))),
      h('section', { class: 'card' }, h('h2', { id: 'quizzes' }, 'Monday quizzes'), h('p', { class: 'muted' }, 'Each quiz covers the previous week. No quiz in week 1 or on the Monday after an exam.'), table(rows(['quiz']))),
      h('section', { class: 'card' }, h('h2', { id: 'exams' }, 'Unit exams (Fridays of weeks 5, 10, 15)'), table(rows(['exam']))),
      custom.length ? h('section', { class: 'card' }, h('h2', {}, 'Teacher-made quizzes'), table(custom)) : null,
      h('p', {}, h('a', { href: '#/templates' }, 'Quiz templates: blank workbooks and the quiz builder →')),
    ]);
  }

  // ===========================================================================
  // Pages: Quiz templates and builder
  // ===========================================================================
  function pageTemplates() {
    const reg = [...registry().values()].filter((e) => e.kind !== 'practice' && e.a.workbook);
    layout('templates', [
      h('h1', {}, 'Quiz templates'),
      h('p', { class: 'lead' }, 'Blank workbooks for every quiz, exam and assignment, plus a builder for writing your own auto-graded quizzes.'),
      h('section', { class: 'card' },
        h('h2', {}, 'Blank workbooks (.xlsx)'),
        h('p', { class: 'muted' }, 'Students using desktop Excel download the template, complete it, and upload it when they submit. It is graded exactly like the browser workbook.'),
        h('div', { class: 'template-grid' }, reg.sort((x, y) => (x.a.week || x.lesson.week) - (y.a.week || y.lesson.week)).map((e) => h('div', { class: 'template-item' },
          badge(kindLabel[e.kind], e.kind), h('span', {}, e.a.title),
          h('button', { class: 'btn btn-sm', type: 'button', onClick: () => X.download(X.write(G.buildTemplate(e.a), { title: e.a.title }), safeFile(e.a.title) + '.xlsx') }, '⬇ .xlsx'))))),
      h('section', { class: 'card' },
        h('h2', {}, 'Write your own quiz'),
        h('p', {}, 'The quiz builder makes the same kind of quiz the course uses: multiple choice, numeric, short-answer and matching questions, plus workbook tasks graded from cell values and formulas. Quizzes you build are saved in this browser, can be taken right away, and export as a file you can share or add to the course.'),
        h('div', { class: 'btn-row' },
          h('a', { class: 'btn btn-primary', href: '#/builder' }, 'Open the quiz builder'),
          fileButton('Import quiz (.json)', '.json', async (file) => {
            try {
              const q = JSON.parse(await readFile(file));
              if (!q || !Array.isArray(q.questions)) throw new Error();
              q.id = 'custom-' + Date.now().toString(36);
              state.custom.push(q); save(true); toast('Imported "' + q.title + '".'); route();
            } catch (e) { toast('That file is not a quiz export.', 'error'); }
          }),
          h('a', { class: 'btn', href: 'docs/AUTHORING.md', target: '_blank', rel: 'noopener' }, 'Authoring guide')),
        state.custom.length ? h('table', { class: 'simple' }, h('thead', {}, h('tr', {}, h('th', {}, 'Your quizzes'), h('th', {}, 'Items'), h('th', {}, ''))),
          h('tbody', {}, state.custom.map((q) => h('tr', {}, h('td', {}, q.title), h('td', {}, String(q.questions.length)),
            h('td', {}, h('a', { class: 'btn btn-sm', href: '#/quiz/' + q.id }, 'Take'), ' ', h('a', { class: 'btn btn-sm', href: '#/builder/' + q.id }, 'Edit'), ' ',
              h('button', { class: 'btn btn-sm', type: 'button', onClick: () => downloadText(JSON.stringify(q, null, 2), safeFile(q.title) + '.quiz.json', 'application/json') }, 'Export'), ' ',
              h('button', { class: 'btn btn-sm btn-danger', type: 'button', onClick: () => dialog('Delete quiz?', h('p', {}, 'Delete "' + q.title + '" from this browser?'), [{ label: 'Cancel' }, { label: 'Delete', primary: true, onClick: () => { state.custom = state.custom.filter((x) => x.id !== q.id); save(true); route(); } }]) }, 'Delete')))))) : null),
      h('section', { class: 'card' },
        h('h2', {}, 'Template for course authors'),
        h('p', {}, 'To add a quiz to the course itself, copy ', h('code', {}, 'assets/js/content/quiz-template.js'), ', fill it in, and load it in index.html. Every question type and option is documented in the file and in the authoring guide. The test suite then proves your quiz can be solved for 100%.'),
        h('a', { class: 'btn', href: 'assets/js/content/quiz-template.js', download: 'quiz-template.js' }, '⬇ quiz-template.js')),
    ]);
  }

  function pageBuilder(editId) {
    const existing = editId ? state.custom.find((q) => q.id === editId) : null;
    const quiz = existing ? JSON.parse(JSON.stringify(existing)) : { id: 'custom-' + Date.now().toString(36), title: 'New quiz', week: currentWeek() || 1, minutes: 15, kind: 'quiz', intro: '', tools: 'basic', workbook: null, questions: [] };
    let builderWb = quiz.workbook ? W.Workbook.fromSpec(quiz.workbook) : null;
    if (quiz.workbookJSON) builderWb = W.Workbook.fromJSON(quiz.workbookJSON);
    let view = null;

    const title = h('input', { type: 'text', value: quiz.title });
    const week = select(C.lessons.map((l) => [l.week, 'Week ' + l.week + ': ' + l.title]), quiz.week);
    const minutes = h('input', { type: 'number', min: 1, max: 180, value: quiz.minutes });
    const intro = h('textarea', { rows: 2 }, quiz.intro || '');
    const tools = select([['basic', 'Basic (formulas, formats)'], ['data', 'Data (sort, split, import)'], ['organize', 'Organize (tables, validation)'], ['analyze', 'Analyze (charts, pivots)'], ['all', 'Everything (incl. Power Query)']], quiz.tools || 'basic');
    const qHost = h('ol', { class: 'builder-questions' });
    const wbHost = h('div', { class: 'wb-host' });

    function renderWb() {
      if (view) { view.destroy(); view = null; }
      wbHost.replaceChildren();
      if (!builderWb) { wbHost.append(h('p', { class: 'muted' }, 'No workbook. Add one to write workbook tasks (cell answers checked automatically).')); return; }
      view = SheetView.mount(wbHost, builderWb, { tools: 'all', fileName: quiz.title, onChange: (w) => { builderWb = w; } });
    }

    function renderQs() {
      qHost.replaceChildren(...quiz.questions.map((q, i) => h('li', { class: 'builder-q' },
        h('div', { class: 'bq-head' }, badge(typeLabel(q.type)), h('strong', {}, q.prompt || '(no prompt)'), h('span', { class: 'muted' }, ' · ' + plural(q.points ?? 1, 'pt')),
          h('span', { class: 'bq-actions' },
            h('button', { class: 'icon-btn', type: 'button', title: 'Move up', 'aria-label': 'Move up', disabled: i === 0, onClick: () => { quiz.questions.splice(i - 1, 0, quiz.questions.splice(i, 1)[0]); renderQs(); } }, '↑'),
            h('button', { class: 'icon-btn', type: 'button', title: 'Edit', 'aria-label': 'Edit', onClick: () => questionDialog(q, (nq) => { quiz.questions[i] = nq; renderQs(); }) }, '✎'),
            h('button', { class: 'icon-btn', type: 'button', title: 'Delete', 'aria-label': 'Delete', onClick: () => { quiz.questions.splice(i, 1); renderQs(); } }, '×'))),
        h('div', { class: 'small muted' }, describeQ(q)))));
      if (!quiz.questions.length) qHost.append(h('li', { class: 'muted' }, 'No questions yet.'));
    }

    function collect() {
      quiz.title = title.value.trim() || 'Untitled quiz';
      quiz.week = Number(week.value);
      quiz.minutes = Math.max(1, Number(minutes.value) || 15);
      quiz.intro = intro.value;
      quiz.tools = tools.value;
      if (builderWb) {
        builderWb.recalc();
        quiz.workbookJSON = builderWb.toJSON();
        quiz.workbook = workbookToSpec(builderWb);
      } else { quiz.workbook = null; delete quiz.workbookJSON; }
      return quiz;
    }

    function validateQuiz(q) {
      const problems = [];
      if (!q.questions.length) problems.push('Add at least one question.');
      q.questions.forEach((x, i) => {
        if (G.isWorkbookQuestion(x) && !q.workbook) problems.push('Question ' + (i + 1) + ' needs a workbook.');
        if (x.type === 'cell' && q.workbook && !q.workbook.sheets.some((s) => s.name === x.sheet)) problems.push('Question ' + (i + 1) + ': sheet "' + x.sheet + '" is not in the workbook.');
      });
      if (!problems.length) {
        // prove it can be solved for 100%
        try {
          const key = G.buildKey(q);
          const r = G.grade(q, { workbook: key, answers: G.perfectAnswers(q) });
          if (r.percent !== 100) problems.push('The answer key only scores ' + r.percent + '%: ' + r.items.filter((it) => it.status !== 'correct').map((it) => it.feedback).join(' '));
        } catch (e) { problems.push('Could not build the answer key: ' + e.message); }
      }
      return problems;
    }

    layout('templates', [
      crumb(['#/templates', 'Quiz templates'], 'Quiz builder'),
      h('h1', {}, existing ? 'Edit quiz' : 'Quiz builder'),
      h('div', { class: 'card builder-meta' }, h('div', { class: 'grid2' }, field('Title', title), field('Given in', week), field('Time limit (minutes)', minutes), field('Workbook tools', tools)), field('Instructions shown before starting', intro)),
      h('div', { class: 'card' }, h('h2', {}, 'Workbook (optional)'),
        h('p', { class: 'muted' }, 'Type the starting data the way students will see it. Lock given data with the Format menu (or leave it unlocked). Workbook tasks below are graded by comparing a student\'s cell to the answer you define.'),
        h('div', { class: 'btn-row' },
          h('button', { class: 'btn', type: 'button', onClick: () => { builderWb = W.Workbook.fromSpec({ sheets: [{ name: 'Sheet1' }] }); renderWb(); } }, builderWb ? 'Start over with a blank sheet' : 'Add a blank workbook'),
          fileButton('Use an .xlsx as the starting workbook', '.xlsx', async (file) => { try { builderWb = await X.read(await readFile(file, 'buffer')); renderWb(); toast('Loaded ' + file.name); } catch (e) { toast('Could not read that file: ' + e.message, 'error'); } }),
          h('button', { class: 'btn', type: 'button', onClick: () => { builderWb = null; renderWb(); } }, 'Remove workbook')),
        wbHost),
      h('div', { class: 'card' }, h('h2', {}, 'Questions'), qHost,
        h('div', { class: 'btn-row' }, ['choice', 'number', 'text', 'match', 'cell'].map((t) => h('button', { class: 'btn', type: 'button', onClick: () => questionDialog({ type: t, id: 'q' + (quiz.questions.length + 1) + '-' + Math.random().toString(36).slice(2, 5), points: 1 }, (nq) => { quiz.questions.push(nq); renderQs(); }, true) }, '+ ' + typeLabel(t))))),
      h('div', { class: 'btn-row sticky-actions' },
        h('button', { class: 'btn btn-primary', type: 'button', onClick: () => {
          const q = collect();
          const problems = validateQuiz(q);
          if (problems.length) { dialog('Fix these first', h('ul', {}, problems.map((p) => h('li', {}, p)))); return; }
          const i = state.custom.findIndex((x) => x.id === q.id);
          if (i >= 0) state.custom[i] = q; else state.custom.push(q);
          save(true);
          toast('Saved. The answer key checks out at 100%.', 'ok');
          location.hash = '#/templates';
        } }, 'Save quiz'),
        h('button', { class: 'btn', type: 'button', onClick: () => { const q = collect(); downloadText(JSON.stringify(q, null, 2), safeFile(q.title) + '.quiz.json', 'application/json'); } }, 'Export .json'),
        h('button', { class: 'btn', type: 'button', onClick: () => { const q = collect(); downloadText(quizToJs(q), safeFile(q.title) + '.js', 'text/javascript'); } }, 'Export as course file (.js)'),
        h('a', { class: 'btn', href: '#/templates' }, 'Cancel')),
    ]);
    renderWb();
    renderQs();

    function questionDialog(q, done, isNew) {
      const prompt = h('textarea', { rows: 2 }, q.prompt || '');
      const points = h('input', { type: 'number', min: 0, step: 0.5, value: q.points ?? 1 });
      const body = h('div', {}, field('Question / task', prompt), field('Points', points));
      let read;
      if (q.type === 'choice') {
        const opts = h('textarea', { rows: 4, placeholder: 'One option per line' }, (q.options || []).join('\n'));
        const ans = h('input', { type: 'text', value: [].concat(q.answer ?? 0).map((x) => Number(x) + 1).join(', '), placeholder: 'e.g. 2  or  1, 3 for several' });
        const explain = h('input', { type: 'text', value: q.explain || '' });
        body.append(field('Options', opts, 'One per line.'), field('Correct option number(s)', ans, 'Several numbers make it select-all-that-apply.'), field('Explanation shown after grading (optional)', explain));
        read = () => {
          const options = opts.value.split('\n').map((s) => s.trim()).filter(Boolean);
          const nums = ans.value.split(/[ ,;]+/).filter(Boolean).map((x) => Number(x) - 1).filter((x) => x >= 0 && x < options.length);
          if (options.length < 2 || !nums.length) return 'Add at least two options and the correct number(s).';
          return { options, answer: nums.length > 1 ? nums : nums[0], multi: nums.length > 1 || undefined, explain: explain.value || undefined };
        };
      } else if (q.type === 'number') {
        const ans = h('input', { type: 'text', value: q.answer ?? '' });
        const tol = h('input', { type: 'text', value: q.tolerance ?? '' , placeholder: 'e.g. 0.01' });
        body.append(field('Correct answer', ans), field('Tolerance (optional)', tol));
        read = () => { const p = F.parseNumberLike(ans.value.trim()); if (!p) return 'The answer must be a number.'; return { answer: p.value, tolerance: tol.value ? Number(tol.value) : undefined }; };
      } else if (q.type === 'text') {
        const acc = h('textarea', { rows: 3, placeholder: 'One accepted answer per line' }, (q.accept || []).join('\n'));
        body.append(field('Accepted answers', acc, 'Matching ignores capitals, extra spaces and punctuation.'));
        read = () => { const accept = acc.value.split('\n').map((s) => s.trim()).filter(Boolean); return accept.length ? { accept } : 'Add at least one accepted answer.'; };
      } else if (q.type === 'match') {
        const pairs = h('textarea', { rows: 5, placeholder: 'Term = definition (one per line)' }, (q.pairs || []).map((p) => p.join(' = ')).join('\n'));
        body.append(field('Pairs', pairs, 'One "term = definition" per line.'));
        read = () => { const ps = pairs.value.split('\n').map((l) => l.split('=').map((s) => s.trim())).filter((p) => p.length >= 2 && p[0] && p[1]).map((p) => [p[0], p.slice(1).join('=')]); return ps.length >= 2 ? { pairs: ps } : 'Add at least two pairs.'; };
      } else if (q.type === 'cell') {
        const sheets = builderWb ? builderWb.sheets.map((s) => s.name) : ['Sheet1'];
        const sheet = select(sheets, q.sheet || sheets[0]);
        const cell = h('input', { type: 'text', value: q.cell || '', placeholder: 'e.g. D2' });
        const sol = h('input', { type: 'text', value: q.solution ?? '', placeholder: 'A value, or a formula like =SUM(B2:B9)' });
        const reqF = h('input', { type: 'checkbox', checked: !!q.requireFormula });
        const funcs = h('input', { type: 'text', value: (q.functions || []).join(', '), placeholder: 'e.g. SUMIF, COUNTIF' });
        const tol = h('input', { type: 'text', value: q.tolerance ?? '' });
        const hint = h('input', { type: 'text', value: q.hint || '' });
        body.append(h('div', { class: 'grid2' }, field('Sheet', sheet), field('Cell', cell)),
          field('Answer key', sol, 'The correct answer is computed from this on the starting workbook, so a formula stays correct if you change the data.'),
          h('label', { class: 'check' }, reqF, ' Require a formula (typed values get half credit)'),
          field('Must use functions (optional)', funcs), h('div', { class: 'grid2' }, field('Tolerance (optional)', tol), field('Hint shown when wrong (optional)', hint)));
        read = () => {
          if (!builderWb) return 'Add a workbook first.';
          if (!F.parseAddress(cell.value.trim().toUpperCase())) return 'Type a cell address like D2.';
          if (!sol.value.trim()) return 'Type the answer or a formula.';
          if (sol.value.trim().startsWith('=')) { try { F.parse(sol.value.trim()); } catch (e) { return 'The formula has a problem: ' + e.message; } }
          return {
            sheet: sheet.value, cell: cell.value.trim().toUpperCase(), solution: sol.value.trim(), requireFormula: reqF.checked || undefined,
            functions: funcs.value.trim() ? funcs.value.split(/[ ,;]+/).filter(Boolean).map((s) => s.toUpperCase()) : undefined,
            tolerance: tol.value ? Number(tol.value) : undefined, hint: hint.value || undefined,
          };
        };
      }
      dialog((isNew ? 'Add ' : 'Edit ') + typeLabel(q.type).toLowerCase(), body, [
        { label: 'Cancel' },
        { label: isNew ? 'Add' : 'Save', primary: true, onClick: () => {
          if (!prompt.value.trim()) { toast('Write the question first.', 'warn'); return false; }
          const extra = read();
          if (typeof extra === 'string') { toast(extra, 'warn'); return false; }
          const nq = { id: q.id, type: q.type, prompt: prompt.value.trim(), points: Number(points.value) || 0, ...extra };
          Object.keys(nq).forEach((k) => nq[k] === undefined && delete nq[k]);
          done(nq);
        } },
      ], { wide: true });
    }
  }

  const typeLabel = (t) => ({ choice: 'Multiple choice', number: 'Number', text: 'Short answer', match: 'Matching', cell: 'Workbook cell' }[t] || t);
  function describeQ(q) {
    if (q.type === 'choice') return q.options.map((o, i) => ([].concat(q.answer).map(Number).includes(i) ? '✓ ' : '· ') + o).join('   ');
    if (q.type === 'number') return 'Answer: ' + q.answer + (q.tolerance ? ' ± ' + q.tolerance : '');
    if (q.type === 'text') return 'Accepts: ' + q.accept.join(' / ');
    if (q.type === 'match') return q.pairs.map((p) => p[0] + ' → ' + p[1]).join('; ');
    if (q.type === 'cell') return q.sheet + '!' + q.cell + ' = ' + q.solution + (q.requireFormula ? ' (formula required)' : '') + (q.functions ? ' using ' + q.functions.join(', ') : '');
    return '';
  }

  // Convert a live workbook into a compact spec that fromSpec can rebuild.
  function workbookToSpec(wb) {
    return {
      sheets: wb.sheets.map((s) => {
        const cells = {};
        for (const [k, c] of Object.entries(s.cells)) {
          if ((c.input === '' || c.input === undefined) && !c.format && !c.style && !c.locked) continue;
          const o = { input: c.input };
          if (c.format) o.format = c.format;
          if (c.style) o.style = c.style;
          if (c.locked) o.locked = true;
          cells[k] = o;
        }
        return { name: s.name, cells, widths: s.widths, validations: s.validations, condFormats: s.condFormats, headerRow: false };
      }),
      tables: wb.tables, names: wb.names, sources: wb.sources,
    };
  }

  function quizToJs(q) {
    const clean = { ...q };
    delete clean.workbookJSON;
    return '/*\n * ' + q.title + ' — exported from the quiz builder.\n * Load this file in index.html after assets/js/content/quizzes.js.\n */\n' +
      '(function (root) {\n  \'use strict\';\n  const C = typeof module === \'object\' && module.exports ? require(\'./course.js\') : root.DIM.Content;\n' +
      '  C.quizzes.push(' + JSON.stringify({ ...clean, id: 'quiz-custom-' + safeFile(q.title).toLowerCase() }, null, 2).replace(/\n/g, '\n  ') + ');\n' +
      '  if (typeof module === \'object\' && module.exports) module.exports = C;\n})(typeof globalThis !== \'undefined\' ? globalThis : this);\n';
  }

  // ===========================================================================
  // Pages: Grades (student view + teacher gradebook)
  // ===========================================================================
  function pageGrades() {
    const reg = [...registry().values()].filter((e) => e.kind !== 'practice' && !e.custom);
    const g = courseGrade((id) => countedSubmission(id));
    const label = { assignment: 'Weekly assignments', quiz: 'Monday quizzes (lowest dropped)', exam1: 'Exam 1', exam2: 'Exam 2', exam3: 'Exam 3', project: 'Final dashboard project' };
    const teacher = state.teacher.receipts;
    layout('grades', [
      h('h1', {}, 'Grades'),
      h('div', { class: 'grid2' },
        h('section', { class: 'card' },
          h('h2', {}, state.student ? state.student.name : 'Your grades'),
          g.overall === null ? h('p', { class: 'muted' }, 'Nothing submitted yet from this browser.') : h('p', { class: 'grade-big' }, pct(g.overall), ' ', h('span', {}, G.letterGrade(g.overall))),
          h('table', { class: 'simple' }, h('tbody', {}, Object.entries(label).map(([k, l]) => h('tr', {}, h('td', {}, l), h('td', { class: 'num' }, g.weights[k] + '%'), h('td', { class: 'num' }, g.parts[k] === null ? '—' : pct(g.parts[k])))))),
          h('p', { class: 'muted small' }, 'Uses auto-graded points. Rubric points appear once your teacher scores them. Categories with no work yet are left out of the average.'),
          h('div', { class: 'btn-row' }, h('button', { class: 'btn', type: 'button', disabled: !state.submissions.length, onClick: () => downloadText(JSON.stringify({ type: 'dim-receipts', receipts: state.submissions.map(receiptOf) }), safeFile(state.student ? state.student.name : 'student') + '_all_receipts.dim.json', 'application/json') }, 'Download all my receipts'))),
        h('section', { class: 'card' }, h('h2', {}, 'Every graded item'),
          h('table', { class: 'simple compact' }, h('thead', {}, h('tr', {}, h('th', {}, 'Item'), h('th', { class: 'num' }, 'Score'))),
            h('tbody', {}, reg.sort((x, y) => (x.a.week || x.lesson.week) - (y.a.week || y.lesson.week)).map((e) => {
              const s = countedSubmission(e.id);
              return h('tr', {}, h('td', {}, h('a', { href: hrefFor(e) }, e.a.title)), h('td', { class: 'num' }, s ? s.score + '/' + s.autoMax + (s.manualMax ? ' +' + (s.manualScore ?? '?') + '/' + s.manualMax : '') : '—'));
            }))))),
      h('section', { class: 'card teacher' },
        h('h2', {}, 'Teacher gradebook'),
        h('p', {}, 'Collect students\' receipt files (.dim.json) and import them here. Each receipt is ', h('strong', {}, 're-graded from the work inside it'), ', so an edited score is caught. Then score rubric points and export a CSV for your gradebook.'),
        h('div', { class: 'btn-row' },
          fileButton('Import receipts', '.json', async (files) => {
            let added = 0, flagged = 0, bad = 0;
            for (const f of files) {
              try {
                const j = JSON.parse(await readFile(f));
                const list = j.type === 'dim-receipts' ? j.receipts : [j];
                for (const r of list) {
                  if (!r || r.type !== 'dim-receipt') { bad++; continue; }
                  const checked = regrade(r);
                  if (checked.mismatch) flagged++;
                  const i = teacher.findIndex((x) => x.id === r.id);
                  if (i >= 0) teacher[i] = { ...teacher[i], ...checked }; else teacher.push(checked);
                  added++;
                }
              } catch (e) { bad++; }
            }
            save(true);
            toast('Imported ' + added + ' receipt(s).' + (flagged ? ' ' + flagged + ' flagged: score did not match the work.' : '') + (bad ? ' ' + bad + ' file(s) skipped.' : ''), flagged || bad ? 'warn' : 'ok');
            route();
          }, { multiple: true, primary: true }),
          h('button', { class: 'btn', type: 'button', disabled: !teacher.length, onClick: () => downloadText(gradebookCSV(teacher), 'gradebook-' + isoDate(new Date()) + '.csv', 'text/csv') }, 'Export gradebook CSV'),
          h('button', { class: 'btn', type: 'button', disabled: !state.submissions.length, title: 'Add this browser\'s own submissions (e.g. a shared classroom computer)', onClick: () => { state.submissions.forEach((s) => { const r = regrade(receiptOf(s)); if (!teacher.some((x) => x.id === r.id)) teacher.push(r); }); save(true); route(); } }, 'Add this browser\'s submissions'),
          h('button', { class: 'btn btn-danger', type: 'button', disabled: !teacher.length, onClick: () => dialog('Clear gradebook?', h('p', {}, 'Remove all imported receipts from this browser?'), [{ label: 'Cancel' }, { label: 'Clear', primary: true, onClick: () => { state.teacher.receipts = []; save(true); route(); } }]) }, 'Clear')),
        teacher.length ? gradebookTable(teacher) : h('p', { class: 'muted' }, 'No receipts imported yet.')),
    ]);
  }

  function regrade(r) {
    const e = find(r.assessmentId);
    const out = { ...r, workbook: undefined, verified: false, mismatch: false, claimed: r.claimed ?? r.score };
    out.hasWorkbook = !!r.workbook;
    if (!e) return { ...out, note: 'Unknown assessment' };
    try {
      const wb = r.workbook ? W.Workbook.fromJSON(r.workbook) : undefined;
      if (wb && e.a.workbook && e.a.workbook.sources) wb.sources = { ...e.a.workbook.sources, ...(wb.sources || {}) };
      const rep = G.grade(e.a, { workbook: wb, answers: r.answers || {} });
      out.verified = true;
      out.regradeScore = rep.score;
      out.mismatch = Math.abs(rep.score - out.claimed) > 0.01;
      out.score = rep.score; out.autoMax = rep.max; out.autoPercent = rep.percent;
      const manualMax = e.a.manualPoints || 0;
      out.max = rep.max + manualMax; out.manualMax = manualMax;
      out.percent = manualMax ? Math.round(((rep.score + (out.manualScore || 0)) / out.max) * 1000) / 10 : rep.percent;
    } catch (err) { out.note = 'Could not re-grade: ' + err.message; }
    // keep the uploaded file and work out of the gradebook table, but keep the file for download
    return out;
  }

  function gradebookTable(list) {
    const students = [...new Set(list.map((r) => r.student.name))].sort();
    const items = [...new Set(list.map((r) => r.assessmentId))].sort((a, b) => sortKey(a) - sortKey(b));
    const pickR = (name, id) => {
      const rs = list.filter((r) => r.student.name === name && r.assessmentId === id).sort((a, b) => a.submittedAt.localeCompare(b.submittedAt));
      if (!rs.length) return null;
      return state.settings.policy === 'best' ? rs.reduce((a, b) => (b.percent > a.percent ? b : a)) : rs[0];
    };
    return h('div', { class: 'table-scroll' }, h('table', { class: 'simple gradebook' },
      h('thead', {}, h('tr', {}, h('th', {}, 'Student'), items.map((id) => h('th', { class: 'num', title: (find(id) || { a: { title: id } }).a.title }, shortId(id))), h('th', { class: 'num' }, 'Course'))),
      h('tbody', {}, students.map((name) => {
        const g = courseGrade((id) => pickR(name, id));
        return h('tr', {}, h('th', { scope: 'row' }, name), items.map((id) => {
          const r = pickR(name, id);
          if (!r) return h('td', { class: 'num muted' }, '—');
          return h('td', { class: 'num' + (r.mismatch ? ' flag' : ''), title: r.mismatch ? 'The receipt claimed ' + r.claimed + ' but the work scores ' + r.regradeScore : r.title },
            h('button', { class: 'linkish', type: 'button', onClick: () => teacherDetail(r) }, r.score + (r.manualMax ? '+' + (r.manualScore ?? '?') : '') + '/' + r.max), r.mismatch ? ' ⚠' : '');
        }), h('td', { class: 'num' }, g.overall === null ? '—' : pct(g.overall) + ' ' + G.letterGrade(g.overall)));
      }))));
  }
  const sortKey = (id) => { const e = find(id); return e ? (e.a.week || e.lesson.week) * 10 + ({ quiz: 0, assignment: 1, project: 1, exam: 2 }[e.kind] || 0) : 999; };
  const shortId = (id) => id.replace('quiz-', 'Q').replace('hw-', 'HW').replace('exam-', 'Exam ');

  function teacherDetail(r) {
    const manual = h('input', { type: 'number', min: 0, max: r.manualMax || 0, step: 0.5, value: r.manualScore ?? '' });
    const e = find(r.assessmentId);
    dialog(r.student.name + ' · ' + r.title, h('div', {},
      h('p', {}, 'Submitted ' + new Date(r.submittedAt).toLocaleString() + (r.auto ? ' (auto-submitted when time ran out)' : '') + ' · ' + (r.mode === 'upload' ? 'uploaded ' + (r.uploadName || '.xlsx') : r.hasWorkbook ? 'web workbook' : 'page answers')),
      r.mismatch ? h('p', { class: 'warn' }, '⚠ The receipt said ' + r.claimed + ' points; the work inside it scores ' + r.regradeScore + '. The re-graded score is used.') : h('p', { class: 'status-ok' }, '✓ Re-graded from the work: ' + r.score + '/' + r.autoMax + ' auto points.'),
      r.manualMax ? h('div', {}, e && e.a.rubric ? rubricCard(e.a) : null, field('Rubric points (out of ' + r.manualMax + ')', manual)) : null,
      r.file ? h('p', {}, 'Attached file: ' + r.file.name + ' ', r.file.data ? h('a', { href: r.file.data, download: r.file.name }, 'Download') : h('span', { class: 'muted' }, '(too large to include; ask the student)')) : null,
      h('ol', { class: 'report-list' }, (r.items || []).map((it) => h('li', { class: it.status }, h('div', { class: 'small' }, it.id + ': ' + it.earned + '/' + it.points + ' · ' + it.feedback))))), [
      { label: 'Close' },
      r.manualMax ? { label: 'Save rubric points', primary: true, onClick: () => {
        const v = manual.value === '' ? null : Math.max(0, Math.min(r.manualMax, Number(manual.value)));
        const t = state.teacher.receipts.find((x) => x.id === r.id);
        t.manualScore = v;
        t.percent = Math.round(((t.score + (v || 0)) / t.max) * 1000) / 10;
        save(true); route();
      } } : null,
    ].filter(Boolean), { wide: true });
  }

  function gradebookCSV(list) {
    const esc = (v) => { const s = String(v ?? ''); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const rows = [['Student', 'Period', 'Item ID', 'Title', 'Kind', 'Submitted', 'Auto score', 'Auto max', 'Rubric score', 'Rubric max', 'Total', 'Max', 'Percent', 'Verified', 'Flag']];
    list.slice().sort((a, b) => a.student.name.localeCompare(b.student.name) || sortKey(a.assessmentId) - sortKey(b.assessmentId) || a.submittedAt.localeCompare(b.submittedAt)).forEach((r) => {
      const total = r.score + (r.manualScore || 0);
      rows.push([r.student.name, r.student.period || '', r.assessmentId, r.title, r.kind, r.submittedAt, r.score, r.autoMax, r.manualScore ?? '', r.manualMax || '', total, r.max, r.percent, r.verified ? 'yes' : 'no', r.mismatch ? 'score did not match work' : '']);
    });
    return rows.map((r) => r.map(esc).join(',')).join('\r\n');
  }

  // ===========================================================================
  // Pages: Help, 404
  // ===========================================================================
  function pageHelp() {
    layout('', h('div', { class: 'narrow prose' },
      crumb(['#/', 'Home'], 'How this site works'),
      h('h1', {}, 'How this site works'),
      h('h2', {}, 'The workbooks'),
      h('p', {}, 'Every workbook on this site behaves like Excel: click a cell and type, start formulas with =, click cells while typing a formula to insert their address, press F4 to add $ signs, and use Ctrl+C, Ctrl+V, Ctrl+Z, Ctrl+D as usual. The ribbon above each workbook only shows the tools that task needs. Download any workbook as a real .xlsx to keep working in Excel.'),
      h('h2', {}, 'Checking and submitting'),
      h('p', {}, 'Practice workbooks have a Check my work button: use it as often as you like. Assignments, quizzes and exams are graded when you click Submit. Feedback appears next to every task and question.'),
      h('h2', {}, 'Using desktop Excel instead'),
      h('p', {}, 'On any assignment, quiz or exam with a workbook, download the blank template, complete it in Excel, then choose "Upload from Excel" and pick your file before you submit. Keep the sheet names the same.'),
      h('h2', {}, 'Where your work is saved'),
      h('p', {}, 'Your progress is saved in this browser automatically. It is not uploaded anywhere. If you switch computers, download a backup in Settings or download your workbook. To hand work in, download the receipt after you submit and turn it in the way your teacher asks.'),
      h('h2', {}, 'For teachers'),
      h('p', {}, 'Set the first Monday of the semester in Settings so due dates line up. Collect receipt files and import them on the Grades page: every receipt is re-graded from the work inside it. Use Quiz Templates to download blank workbooks or build your own auto-graded quizzes.'),
    ));
  }

  function notFound() {
    layout('', h('div', { class: 'narrow' }, h('h1', {}, 'Page not found'), h('p', {}, h('a', { href: '#/' }, 'Go home'))));
  }

  // ===========================================================================
  // Router
  // ===========================================================================
  function route() {
    clearInterval(timerHandle);
    document.querySelectorAll('.modal-overlay').forEach((m) => m.remove());
    const hash = location.hash.replace(/^#\/?/, '');
    const [page, a, b] = hash.split('/');
    const titles = { '': 'Home', lessons: 'Lessons', work: 'Assignments & Quizzes', templates: 'Quiz Templates', grades: 'Grades', settings: 'Settings', help: 'Help', builder: 'Quiz builder' };
    let title = titles[page] || '';
    switch (page) {
      case '': pageHome(); break;
      case 'lessons': pageLessons(); break;
      case 'week': pageWeek(Number(a), b); title = 'Week ' + a + (lessonOf(Number(a)) ? ': ' + lessonOf(Number(a)).title : ''); break;
      case 'assignment': pageAssignment(a); title = (find(a) || { a: { title: '' } }).a.title; break;
      case 'quiz': case 'exam': pageQuiz(a); title = (find(a) || { a: { title: '' } }).a.title; break;
      case 'work': pageWork(); break;
      case 'templates': pageTemplates(); break;
      case 'builder': pageBuilder(a); break;
      case 'grades': pageGrades(); break;
      case 'settings': pageSettings(); break;
      case 'help': pageHelp(); break;
      default: notFound();
    }
    document.title = (title ? title + ' · ' : '') + C.course.short;
    if (!(page === 'week' && b)) window.scrollTo(0, 0);
  }

  addEventListener('hashchange', route);
  route();

  // exposed for tests and the console
  root.DIM.App = { state: () => state, registry, courseGrade, regrade, gradebookCSV, workbookToSpec, route };
})(typeof globalThis !== 'undefined' ? globalThis : this);
