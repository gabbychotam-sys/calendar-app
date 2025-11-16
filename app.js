let rawEvents = [];
let currentView = 'cards';

const excelInput = document.getElementById('excelFile');
const statusEl = document.getElementById('statusMessage');
const eventsContainer = document.getElementById('eventsContainer');
const cardsBtn = document.getElementById('cardsViewBtn');
const tableBtn = document.getElementById('tableViewBtn');
const exitFullscreenBtn = document.getElementById('exitFullscreenBtn');

function setStatus(msg, type = '') {
  statusEl.textContent = msg;
  statusEl.classList.remove('status-ok', 'status-error');
  if (type === 'ok') statusEl.classList.add('status-ok');
  if (type === 'error') statusEl.classList.add('status-error');
}

function excelToDate(value) {
  if (value == null) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  if (typeof value === 'number' && typeof XLSX !== 'undefined' && XLSX.SSF) {
    const d = XLSX.SSF.parse_date_code(value);
    if (!d) return null;
    return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
  }
  const s = String(value).replace('T', ' ');
  const d = new Date(s);
  if (!isNaN(d)) return d;
  return null;
}

// load from localStorage on startup
window.addEventListener('DOMContentLoaded', () => {
  const saved = localStorage.getItem('calendarEventsV1');
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      rawEvents = parsed.map(ev => ({
        subject: ev.subject,
        location: ev.location,
        start: new Date(ev.start),
        end: ev.end ? new Date(ev.end) : null
      }));
      setStatus('האירועים נטענו מהזיכרון (שמירה אוטומטית).', 'ok');
      showToday();
    } catch (e) {
      console.error(e);
    }
  }
});

function saveEvents() {
  const toSave = rawEvents.map(ev => ({
    subject: ev.subject,
    location: ev.location,
    start: ev.start.toISOString(),
    end: ev.end ? ev.end.toISOString() : null
  }));
  localStorage.setItem('calendarEventsV1', JSON.stringify(toSave));
}

excelInput.addEventListener('change', handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet);

      rawEvents = rows.map(r => {
        const start = excelToDate(r['התחלה']);
        const end = excelToDate(r['סיום']);
        return {
          subject: r['נושא'] || '',
          location: r['מיקום'] || '',
          start,
          end
        };
      }).filter(ev => ev.start);

      if (!rawEvents.length) {
        setStatus('לא נמצאו אירועים בקובץ או שפורמט העמודות שונה.', 'error');
        eventsContainer.innerHTML = '<p class="no-events">אין אירועים להצגה.</p>';
        return;
      }

      saveEvents();
      setStatus(`הקובץ נטען בהצלחה. נמצאו ${rawEvents.length} אירועים.`, 'ok');
      showToday();
    } catch (err) {
      console.error(err);
      setStatus('שגיאה בטעינת הקובץ. ודא שהקובץ סגור באקסל ושהעמודות הן: נושא, מיקום, התחלה, סיום.', 'error');
    }
  };

  reader.onerror = function() {
    setStatus('שגיאה בקריאת הקובץ. נסה שוב.', 'error');
  };

  reader.readAsArrayBuffer(file);
}

function startOfDay(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function endOfDay(d) {
  const x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
}

function showToday() {
  const today = new Date();
  showEventsInRange(startOfDay(today), endOfDay(today));
}

function formatDate(d) {
  return d.toLocaleDateString('he-IL', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
}

function formatTime(d) {
  return d.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' });
}

function showEventsInRange(from, to) {
  if (!rawEvents.length) {
    eventsContainer.innerHTML = '<p class="hint">עדיין לא נטען קובץ אקסל.</p>';
    return;
  }
  const list = rawEvents
    .filter(ev => ev.start >= from && ev.start <= to)
    .sort((a, b) => a.start - b.start);

  if (currentView === 'cards') {
    renderCards(list, from, to);
  } else {
    renderTable(list, from, to);
  }
}

function renderCards(list, from, to) {
  if (!list.length) {
    eventsContainer.innerHTML = '<p class="no-events">אין אירועים בטווח שנבחר.</p>';
    return;
  }
  let html = '';
  html += `<p class="hint">מציג ${list.length} אירועים בין ${formatDate(from)} לבין ${formatDate(to)}.</p>`;
  list.forEach(ev => {
    html += `
      <article class="event-card">
        <h3>${ev.subject || 'ללא נושא'}</h3>
        <div class="event-meta">
          <span>📅 ${formatDate(ev.start)}</span>
          <span>⏰ ${formatTime(ev.start)} - ${formatTime(ev.end || ev.start)}</span>
          ${ev.location ? `<span>📍 ${ev.location}</span>` : ''}
        </div>
      </article>
    `;
  });
  eventsContainer.innerHTML = html;
}

function renderTable(list, from, to) {
  if (!list.length) {
    eventsContainer.innerHTML = '<p class="no-events">אין אירועים בטווח שנבחר.</p>';
    return;
  }
  let html = '';
  html += `<p class="hint">מציג ${list.length} אירועים בין ${formatDate(from)} לבין ${formatDate(to)}.</p>`;
  html += '<table class="events-table"><thead><tr><th>נושא</th><th>תאריך</th><th>שעות</th><th>מיקום</th></tr></thead><tbody>';
  list.forEach(ev => {
    html += `
      <tr>
        <td>${ev.subject || 'ללא נושא'}</td>
        <td>${formatDate(ev.start)}</td>
        <td>${formatTime(ev.start)} - ${formatTime(ev.end || ev.start)}</td>
        <td>${ev.location || ''}</td>
      </tr>
    `;
  });
  html += '</tbody></table>';
  eventsContainer.innerHTML = html;
}

// buttons
document.getElementById('todayBtn').onclick = showToday;

document.getElementById('tomorrowBtn').onclick = () => {
  const t = new Date();
  t.setDate(t.getDate() + 1);
  showEventsInRange(startOfDay(t), endOfDay(t));
};

document.getElementById('weekBtn').onclick = () => {
  const from = startOfDay(new Date());
  const to = endOfDay(new Date());
  to.setDate(to.getDate() + 7);
  showEventsInRange(from, to);
};

document.getElementById('monthBtn').onclick = () => {
  const from = startOfDay(new Date());
  const to = endOfDay(new Date());
  to.setMonth(to.getMonth() + 1);
  showEventsInRange(from, to);
};

document.getElementById('showBtn').onclick = () => {
  const dateVal = document.getElementById('dateTo').value;
  if (!dateVal) {
    showToday();
    return;
  }
  const from = startOfDay(new Date());
  const to = endOfDay(new Date(dateVal));
  showEventsInRange(from, to);
};

cardsBtn.onclick = () => {
  currentView = 'cards';
  cardsBtn.classList.add('active');
  tableBtn.classList.remove('active');
  showToday();
};

tableBtn.onclick = () => {
  currentView = 'table';
  tableBtn.classList.add('active');
  cardsBtn.classList.remove('active');
  showToday();
};

// fullscreen & print
document.getElementById('fullscreenBtn').onclick = () => {
  document.body.classList.add('fullscreen-active');
  exitFullscreenBtn.style.display = 'inline-block';
};

exitFullscreenBtn.onclick = () => {
  document.body.classList.remove('fullscreen-active');
  exitFullscreenBtn.style.display = 'none';
};

document.getElementById('printBtn').onclick = () => {
  window.print();
};
