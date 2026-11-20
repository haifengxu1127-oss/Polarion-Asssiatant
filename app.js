/* All spreadsheet and folder processing happens in this browser tab. */
const $ = id => document.getElementById(id);
const rows = [];
const fileUrls = new Map();
const priorities = ['', '0', '30', '50', '70', '100'];
const supported = /\.(pdf|docx|txt|md)$/i;
let nextId = 1;

function setStatus(message, error = false) {
  $('status').textContent = message;
  $('status').classList.toggle('error', error);
}

function normalized(value) {
  return String(value ?? '').trim().toLowerCase().replace(/[\s_-]+/g, '');
}

function titleKey(value) {
  return String(value ?? '').trim().toLocaleLowerCase();
}

function isSharePointUrl(value) {
  if (!value) return true;
  try {
    const url = new URL(value);
    return url.protocol === 'https:' && url.hostname.toLowerCase().endsWith('.sharepoint.com');
  } catch { return false; }
}

function createRow(title = '') {
  const row = { id: nextId++, title, hyperlink: '', priority: '', scope: '', file: null, relativePath: '' };
  rows.push(row);
  return row;
}

function uniqueRow(title) {
  const matches = rows.filter(row => titleKey(row.title) === titleKey(title));
  if (matches.length > 1) return { ambiguous: true };
  return { row: matches[0] || createRow(title) };
}

async function readSheet(file) {
  if (!window.XLSX) throw new Error('The spreadsheet reader did not load. Reload the page.');
  const data = await file.arrayBuffer();
  const book = XLSX.read(data, { type: 'array', cellLinks: true });
  const sheet = book.Sheets[book.SheetNames[0]];
  if (!sheet) throw new Error('The workbook has no readable first sheet.');
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  if (!grid.length) throw new Error('The first sheet is empty.');
  const headers = grid[0].map(normalized);
  const column = (...names) => headers.findIndex(header => names.some(name => normalized(name) === header));
  const titleCol = column('File Title', 'File Name', 'Name', 'Title');
  const linkCol = column('Hyperlink', 'SharePoint Link', 'Link', 'URL', 'webUrl');
  const priorityCol = column('Priority');
  const scopeCol = column('Scope', 'Scope(s)', 'Classification');
  if (titleCol < 0) throw new Error('The first sheet needs a File Title column.');
  return grid.slice(1).map((cells, index) => {
    const title = String(cells[titleCol] ?? '').trim();
    const titleCell = sheet[XLSX.utils.encode_cell({ r: index + 1, c: titleCol })];
    const linkCell = linkCol >= 0 ? sheet[XLSX.utils.encode_cell({ r: index + 1, c: linkCol })] : null;
    const hyperlink = String((linkCol >= 0 ? cells[linkCol] : '') || linkCell?.l?.Target || titleCell?.l?.Target || '').trim();
    return { title, hyperlink, priority: priorityCol < 0 ? '' : String(cells[priorityCol] ?? '').trim(), scope: scopeCol < 0 ? '' : String(cells[scopeCol] ?? '').trim() };
  }).filter(item => item.title);
}

function cellInput(td, value, label, onChange) {
  const input = document.createElement('input');
  input.value = value;
  input.setAttribute('aria-label', label);
  input.addEventListener('change', () => onChange(input.value.trim()));
  td.append(input);
  return input;
}

function render() {
  const body = $('rows');
  body.replaceChildren();
  $('export').disabled = rows.length === 0;
  if (!rows.length) {
    const tr = body.insertRow();
    const td = tr.insertCell();
    td.colSpan = 5;
    td.className = 'empty';
    td.textContent = 'No documents yet.';
    return;
  }
  for (const row of rows) {
    const tr = body.insertRow();
    const titleCell = tr.insertCell();
    if (row.file) {
      const link = document.createElement('a');
      link.className = 'title-link';
      link.textContent = row.title;
      link.href = fileUrls.get(row.id);
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      titleCell.append(link);
      if (row.relativePath) {
        const meta = document.createElement('span');
        meta.className = 'row-meta';
        meta.textContent = row.relativePath;
        titleCell.append(meta);
      }
    } else {
      cellInput(titleCell, row.title, 'File Title', value => { row.title = value; });
    }
    const hyperlinkCell = tr.insertCell();
    cellInput(hyperlinkCell, row.hyperlink, `SharePoint Hyperlink for ${row.title}`, value => {
      row.hyperlink = value;
      if (value && !isSharePointUrl(value)) setStatus(`Check the SharePoint URL for ${row.title}.`, true);
    }).placeholder = 'Verified SharePoint URL';
    const priorityCell = tr.insertCell();
    const priority = document.createElement('select');
    priority.setAttribute('aria-label', `Priority for ${row.title}`);
    for (const value of priorities) {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value || 'Unassigned';
      priority.append(option);
    }
    priority.value = priorities.includes(row.priority) ? row.priority : '';
    priority.addEventListener('change', () => { row.priority = priority.value; });
    priorityCell.append(priority);
    const scopeCell = tr.insertCell();
    cellInput(scopeCell, row.scope, `Scopes for ${row.title}`, value => { row.scope = value; }).placeholder = 'e.g. SW, PM';
    const removeCell = tr.insertCell();
    const remove = document.createElement('button');
    remove.className = 'remove';
    remove.type = 'button';
    remove.textContent = '×';
    remove.setAttribute('aria-label', `Remove ${row.title}`);
    remove.addEventListener('click', () => {
      const index = rows.indexOf(row);
      if (index >= 0) rows.splice(index, 1);
      const url = fileUrls.get(row.id);
      if (url) URL.revokeObjectURL(url);
      fileUrls.delete(row.id);
      render();
      setStatus(`${rows.length} rows ready for review.`);
    });
    removeCell.append(remove);
  }
}

async function importSpreadsheet(file, kind) {
  const items = await readSheet(file);
  let updated = 0, ambiguous = 0, invalidLinks = 0;
  for (const item of items) {
    const found = uniqueRow(item.title);
    if (found.ambiguous) { ambiguous++; continue; }
    const row = found.row;
    if (item.hyperlink) {
      if (isSharePointUrl(item.hyperlink)) row.hyperlink = item.hyperlink;
      else invalidLinks++;
    }
    if (kind === 'analysis') {
      if (priorities.includes(item.priority)) row.priority = item.priority;
      if (item.scope) row.scope = item.scope;
    }
    updated++;
  }
  render();
  const problems = [ambiguous && `${ambiguous} duplicate titles need manual review`, invalidLinks && `${invalidLinks} non-SharePoint links skipped`].filter(Boolean);
  setStatus(`Imported ${updated} ${kind === 'links' ? 'link-list' : 'Copilot-result'} rows. ${rows.length} rows ready.${problems.length ? ' ' + problems.join('; ') + '.' : ''}`, problems.length > 0);
}

function wireSpreadsheetInput(id, labelId, kind) {
  $(id).addEventListener('change', async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    $(labelId).textContent = file.name;
    try { await importSpreadsheet(file, kind); }
    catch (error) { setStatus(error.message || 'Could not read the spreadsheet.', true); }
    event.target.value = '';
  });
}

wireSpreadsheetInput('links-file', 'links-name', 'links');
wireSpreadsheetInput('analysis-file', 'analysis-name', 'analysis');

$('folder-file').addEventListener('change', event => {
  const files = [...(event.target.files || [])].filter(file => supported.test(file.name));
  if (!files.length) { setStatus('No supported PDF, Word, TXT, or Markdown files were found.', true); return; }
  let added = 0, matched = 0, ambiguous = 0;
  for (const file of files) {
    const matches = rows.filter(row => titleKey(row.title) === titleKey(file.name));
    const available = matches.filter(row => !row.file);
    const row = matches.length === 1 && available.length === 1 ? available[0] : createRow(file.name);
    if (row === available[0]) matched++;
    if (matches.length > 1) ambiguous++;
    row.file = file;
    row.relativePath = file.webkitRelativePath || file.name;
    fileUrls.set(row.id, URL.createObjectURL(file));
    added++;
  }
  $('folder-name').textContent = `${added} supported files selected`;
  render();
  setStatus(`Selected ${added} local documents; ${matched} matched existing rows.${ambiguous ? ` ${ambiguous} duplicate titles need manual review.` : ''}`, ambiguous > 0);
  event.target.value = '';
});

$('add-row').addEventListener('click', () => {
  createRow();
  render();
  setStatus('Added a blank row. Enter a File Title before export.');
});

$('export').addEventListener('click', () => {
  if (rows.some(row => !row.title.trim())) { setStatus('Every row needs a File Title.', true); return; }
  const invalid = rows.find(row => row.hyperlink && !isSharePointUrl(row.hyperlink));
  if (invalid) { setStatus(`Hyperlink for ${invalid.title} must be an HTTPS SharePoint URL.`, true); return; }
  const table = [['File Title', 'Hyperlink', 'Priority', 'Scope'], ...rows.map(row => [row.title, row.hyperlink, row.priority, row.scope])];
  const sheet = XLSX.utils.aoa_to_sheet(table);
  sheet['!cols'] = [{ wch: 48 }, { wch: 80 }, { wch: 12 }, { wch: 30 }];
  rows.forEach((row, index) => {
    if (row.hyperlink) sheet[`B${index + 2}`].l = { Target: row.hyperlink };
  });
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, 'Requirements');
  XLSX.writeFile(book, 'polarion_requirements.xlsx');
  setStatus(`Exported ${rows.length} rows to polarion_requirements.xlsx.`);
});

window.addEventListener('pagehide', () => {
  for (const url of fileUrls.values()) URL.revokeObjectURL(url);
});

render();
