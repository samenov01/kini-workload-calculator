'use strict';

const $ = id => document.getElementById(id);
const state = {files: new Map(), groups: [], token: '', busy: false, uploading: false, issues: [], limit: 100, hasResults: false};
const format = new Intl.NumberFormat('ru-RU', {maximumFractionDigits: 2});
const programs = [['', 'Выберите'], ['PE', 'PE'], ['CS', 'CS'], ['LOG', 'LOG'], ['MECH', 'MECH'], ['EEE', 'EEE'], ['IM', 'IM'], ['OTHER', 'Другая']];
const languages = [['KZ', 'Казахский'], ['RU', 'Русский'], ['ENG', 'Английский'], ['DE', 'Немецкий']];
function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}
function options(select, entries, selected) {
  select.replaceChildren(...entries.map(([value, label]) => new Option(label, value)));
  select.value = selected ?? '';
}
function error(message) {
  $('error').textContent = message;
  $('error').hidden = !message;
  if (message) $('error').focus();
}
async function api(url, init = {}) {
  const response = await fetch(url, {...init, headers: {'X-KINI-Token': state.token, ...init.headers}});
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || 'Не удалось выполнить запрос.');
  return data;
}
function metadata(group) {
  return state.files.get(group.file_id)?.sheets.find(s => s.sheet === group.sheet);
}
function changed() {
  if (state.hasResults) $('stale').hidden = false;
  try { localStorage.setItem('kini-selection-v1', JSON.stringify(state.groups)); } catch { /* Private mode can disable storage. */ }
  metrics();
}
function metrics() {
  $('metric-files').textContent = new Set(state.groups.map(g => g.file_id)).size;
  $('metric-groups').textContent = state.groups.length;
  $('metric-students').textContent = format.format(state.groups.reduce((sum, g) => sum + (Number(g.students) || 0), 0));
  const years = new Set();
  for (const g of state.groups) {
    const intake = Number(g.intake);
    if (intake >= 2000 && intake <= 2200) {
      for (let i = 0; i < Math.ceil((metadata(g)?.semesters || 0) / 2); i++) years.add(intake + i);
    }
  }
  $('metric-years').textContent = years.size ? `${Math.min(...years)} — ${Math.max(...years) + 1}` : '—';
  $('group-count').textContent = `Групп: ${state.groups.length}`;
  $('run-summary').textContent = `Учебных лет: ${years.size} · файлов Excel: ${years.size}`;
  $('empty').hidden = !!state.groups.length;
  lockControls();
}
function lockControls() {
  const disabled = state.busy || state.uploading;
  document.querySelectorAll('#groups input, #groups select, #groups button, #choose-files, #file-picker, #add-group, #generate').forEach(n => { n.disabled = disabled; });
  $('generate').disabled = disabled || !state.groups.length;
  $('add-group').disabled = disabled || !state.files.size;
  $('generate').textContent = state.busy ? 'Формируем файлы…' : 'Сформировать расчёт →';
  $('calc-form').setAttribute('aria-busy', String(state.busy));
}
function uniqueName(base) {
  const names = new Set(state.groups.map(g => g.name.toLocaleLowerCase()));
  let name = base, n = 2;
  while (names.has(name.toLocaleLowerCase())) name = `${base}-${n++}`;
  return name;
}
function newGroup(file) {
  const sheet = file.sheets[0];
  const intake = sheet.intake || new Date().getFullYear();
  return {file_id: file.id, sheet: sheet.sheet, program: sheet.program, name: uniqueName(`${sheet.program || 'Группа'}-${String(intake).slice(-2)}`), intake, students: '', language: 'KZ'};
}
function field(group, key, title, kind, entries) {
  const label = el('label', '', title);
  const input = el(kind === 'select' ? 'select' : 'input');
  input.name = key;
  input.required = true;
  if (kind === 'select') options(input, entries, group[key]);
  else {
    input.type = kind;
    input.value = group[key] ?? '';
    if (kind === 'number') {
      input.min = key === 'students' ? '1' : '2000';
      input.max = key === 'students' ? '10000' : '2200';
      input.step = '1';
      input.inputMode = 'numeric';
    } else input.maxLength = 60;
  }
  const hint = el('small', 'field-error', 'Проверьте значение поля.');
  input.addEventListener('input', () => {
    input.setCustomValidity('');
    input.removeAttribute('aria-invalid');
    group[key] = kind === 'number' ? (input.value === '' ? '' : Number(input.value)) : input.value;
    changed();
  });
  input.addEventListener('blur', () => input.setAttribute('aria-invalid', String(!input.validity.valid)));
  label.append(input, hint);
  return label;
}
function renderGroups(focusIndex) {
  $('groups').replaceChildren();
  state.groups.forEach((group, index) => {
    const file = state.files.get(group.file_id);
    const meta = metadata(group);
    const box = el('fieldset', 'group');
    box.append(el('legend', 'visually-hidden', `Группа ${index + 1}`));
    const header = el('div', 'group-header');
    const info = el('div', 'file-info');
    info.append(el('strong', '', file.name), el('small', '', `Шаблон распознан · ${meta.semesters} семестров · дисциплин: ${meta.disciplines}`));
    const actions = el('div', 'row-actions');
    for (const [label, icon, action, cls] of [
      ['Создать ещё одну группу с этим МУПом', '⧉', () => { state.groups.splice(index + 1, 0, {...group, name: uniqueName(group.name)}); changed(); renderGroups(index + 1); }, ''],
      ['Убрать группу из расчёта', '×', () => { state.groups.splice(index, 1); changed(); renderGroups(Math.min(index, state.groups.length - 1)); }, ' remove']
    ]) {
      const button = el('button', `icon-button${cls}`, icon);
      button.type = 'button'; button.title = label; button.setAttribute('aria-label', `${label}: ${group.name}`);
      button.addEventListener('click', action); actions.append(button);
    }
    header.append(el('span', 'program-badge', String(index + 1).padStart(2, '0')), info, actions);
    const fields = el('div', 'fields');
    fields.append(field(group, 'name', 'Группа', 'text'), field(group, 'program', 'Программа', 'select', programs),
      field(group, 'intake', 'Год поступления', 'number'), field(group, 'students', 'Студентов', 'number'),
      field(group, 'language', 'Язык обучения', 'select', languages));
    const details = el('details', 'source-select');
    details.append(el('summary', '', `Источник: ${group.sheet} · изменить МУП или лист`));
    const sources = el('div', 'source-fields');
    const fileLabel = el('label', '', 'МУП'), sheetLabel = el('label', '', 'Лист МУПа');
    const fileSelect = el('select'), sheetSelect = el('select');
    options(fileSelect, [...state.files.values()].map(f => [f.id, f.name]), group.file_id);
    options(sheetSelect, file.sheets.map(s => [s.sheet, s.sheet]), group.sheet);
    fileSelect.addEventListener('change', () => {
      const next = state.files.get(fileSelect.value);
      group.file_id = next.id; group.sheet = next.sheets[0].sheet;
      group.program = next.sheets[0].program;
      if (next.sheets[0].intake) group.intake = next.sheets[0].intake;
      changed(); renderGroups(index);
    });
    sheetSelect.addEventListener('change', () => {
      group.sheet = sheetSelect.value;
      const next = metadata(group);
      group.program = next.program;
      if (next.intake) group.intake = next.intake;
      changed(); renderGroups(index);
    });
    fileLabel.append(fileSelect); sheetLabel.append(sheetSelect); sources.append(fileLabel, sheetLabel); details.append(sources);
    box.append(header, fields, details); $('groups').append(box);
  });
  metrics();
  if (focusIndex >= 0) $('groups').children[focusIndex]?.querySelector('input').focus();
}
async function upload(files) {
  if (state.busy || state.uploading || !files.length) return;
  state.uploading = true; error(''); lockControls();
  let added = 0;
  const failures = [];
  for (const [i, file] of [...files].entries()) {
    $('upload-status').textContent = `Читаем ${i + 1} из ${files.length}: ${file.name}`;
    try {
      if (!file.name.toLowerCase().endsWith('.xlsx')) throw new Error('Нужен формат .xlsx.');
      if (file.size > 12 * 1024 * 1024) throw new Error('Файл превышает 12 МБ.');
      const item = await api('/api/upload', {method: 'POST', headers: {'X-Filename': encodeURIComponent(file.name), 'Content-Type': 'application/octet-stream'}, body: file});
      state.files.set(item.id, item);
      if (!state.groups.some(g => g.file_id === item.id)) { state.groups.push(newGroup(item)); added++; }
    } catch (exc) { failures.push(`${file.name}: ${exc.message}`); }
  }
  state.uploading = false; $('file-picker').value = '';
  $('upload-status').textContent = `Добавлено групп: ${added}. Проверьте год поступления и укажите численность. Повторные файлы не дублируются; для второй группы используйте ⧉.`;
  changed(); renderGroups(added ? state.groups.length - 1 : undefined);
  if (failures.length) error(failures.join('\n'));
}
function downloadLink(url, label, className = 'button secondary') {
  const a = el('a', className, label); a.href = url; a.download = ''; return a;
}
function showResults(job) {
  state.hasResults = true; state.issues = job.issues; state.limit = 100;
  $('stale').hidden = true; $('results-empty').hidden = true;
  $('download-all').hidden = false; $('download-all').href = job.archive; $('download-all').download = '';
  $('download-report').href = job.report; $('download-report').download = '';
  $('result-count').textContent = job.results.length;
  $('result-files').replaceChildren();
  for (const result of job.results) {
    const card = el('article', 'result-card'), header = el('header');
    header.append(el('h3', '', `${result.year}–${result.year + 1}`), el('span', 'checked', '✓ Суммы проверены'));
    const numbers = el('div', 'result-numbers');
    for (const [key, title] of [['1', 'I семестр · ч'], ['2', 'II семестр · ч'], ['0', 'За год · ч']]) {
      const item = el('div'); item.append(el('span', '', title), el('strong', '', format.format(result.totals[key]))); numbers.append(item);
    }
    const footer = el('footer'); footer.append(el('small', '', `Строк нагрузки: ${result.load_rows} · лист «Расчет»`), downloadLink(result.url, 'Скачать .xlsx ↓'));
    card.append(header, numbers, footer); $('result-files').append(card);
  }
  $('review').hidden = false;
  $('review-summary').textContent = state.issues.length ? `Замечаний: ${state.issues.length}. Проверьте их перед использованием итогов.` : 'Замечаний к исходным данным нет.';
  options($('issue-kind'), [['', 'Все типы'], ...[...new Set(state.issues.map(i => i.kind))].sort().map(k => [k, k])], '');
  options($('issue-year'), [['', 'Все годы'], ...job.results.map(r => [String(r.year), `${r.year}–${r.year + 1}`])], '');
  $('issue-search').value = ''; filterIssues();
}
function filterIssues() {
  const kind = $('issue-kind').value, year = $('issue-year').value, search = $('issue-search').value.trim().toLocaleLowerCase();
  const filtered = state.issues.filter(i => (!kind || i.kind === kind) && (!year || !i.year || String(i.year) === year) && (!search || `${i.kind} ${i.group} ${i.source} ${i.message}`.toLocaleLowerCase().includes(search)));
  $('issue-rows').replaceChildren();
  for (const issue of filtered.slice(0, state.limit)) {
    const row = el('tr'), source = el('td', '', issue.group || 'Все группы');
    source.append(el('small', '', `${issue.source || ''}${issue.row ? ` · строка ${issue.row}` : ''}${issue.year ? ` · ${issue.year}–${issue.year + 1}` : ''}`));
    row.append(el('td', '', issue.kind), source, el('td', '', issue.message)); $('issue-rows').append(row);
  }
  $('issue-counter').textContent = `Показано ${Math.min(state.limit, filtered.length)} из ${filtered.length}`;
  $('more-issues').hidden = filtered.length <= state.limit;
}
async function watchJob(id) {
  state.busy = true; lockControls(); $('job-progress').hidden = false;
  try {
    for (;;) {
      const job = await api(`/api/jobs/${id}`);
      $('job-message').textContent = job.message;
      if (job.status === 'error') throw new Error(job.message);
      if (job.status === 'complete') {
        showResults(job);
        $('results-title').tabIndex = -1; $('results-title').focus({preventScroll: true});
        $('results').scrollIntoView({behavior: 'smooth', block: 'start'});
        break;
      }
      await new Promise(resolve => setTimeout(resolve, 1200));
    }
  } catch (exc) { error(`${exc.message} Если связь прервалась, обновите страницу: текущий расчёт продолжится.`); }
  finally { state.busy = false; $('job-progress').hidden = true; lockControls(); }
}
$('choose-files').addEventListener('click', () => $('file-picker').click());
$('file-picker').addEventListener('change', event => upload(event.target.files));
for (const type of ['dragenter', 'dragover']) $('dropzone').addEventListener(type, event => { event.preventDefault(); $('dropzone').classList.add('dragover'); });
for (const type of ['dragleave', 'drop']) $('dropzone').addEventListener(type, event => { event.preventDefault(); $('dropzone').classList.remove('dragover'); });
$('dropzone').addEventListener('drop', event => upload(event.dataTransfer.files));
$('add-group').addEventListener('click', () => { state.groups.push(newGroup(state.files.values().next().value)); changed(); renderGroups(state.groups.length - 1); });
$('calc-form').addEventListener('invalid', event => event.target.setAttribute('aria-invalid', 'true'), true);
$('calc-form').addEventListener('submit', async event => {
  event.preventDefault();
  if (state.busy || state.uploading) return;
  error('');
  const names = new Set();
  for (const [index, group] of state.groups.entries()) {
    const name = group.name.trim().toLocaleLowerCase();
    if (!name || names.has(name)) {
      const input = $('groups').children[index].querySelector('input[name=name]');
      input.setCustomValidity('Укажите уникальное название группы.'); input.reportValidity(); return;
    }
    names.add(name);
  }
  if (!$('calc-form').reportValidity()) return;
  state.busy = true; lockControls();
  try {
    const job = await api('/api/generate', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({groups: state.groups})});
    await watchJob(job.id);
  } catch (exc) { state.busy = false; lockControls(); error(exc.message); }
});
for (const id of ['issue-kind', 'issue-year', 'issue-search']) $(id).addEventListener('input', () => { state.limit = 100; filterIssues(); });
$('more-issues').addEventListener('click', () => { state.limit += 100; filterIssues(); });
async function initialize() {
  lockControls();
  try {
    const data = await api('/api/state');
    state.token = data.token;
    state.files = new Map(data.files.map(f => [f.id, f]));
    let selected = data.groups;
    try {
      const saved = JSON.parse(localStorage.getItem('kini-selection-v1'));
      const revision = localStorage.getItem('kini-selection-revision-v1');
      if (Array.isArray(saved) && revision === data.selection_revision) selected = saved;
      else localStorage.setItem('kini-selection-v1', JSON.stringify(selected));
      localStorage.setItem('kini-selection-revision-v1', data.selection_revision);
    } catch { /* Saved server selection remains available. */ }
    state.groups = selected.filter(g => state.files.get(g.file_id)?.sheets.some(s => s.sheet === g.sheet));
    renderGroups();
    const completed = data.jobs.filter(j => j.status === 'complete').at(-1);
    if (completed) { showResults(completed); $('stale').hidden = false; }
    const active = data.jobs.find(j => ['queued', 'running'].includes(j.status));
    if (active) await watchJob(active.id);
  } catch (exc) { error(`Не удалось загрузить приложение: ${exc.message}`); }
  finally { $('loading').hidden = true; }
}
initialize();
