const $ = (s) => document.querySelector(s);
const form = $('#task-form');
const taskList = $('#task-list');
const emptyState = $('#empty-state');
const toastNode = $('#toast');
let tasks = [];
let activeFilter = 'active';
let audioContext;
const alarmTimers = new Map();
let pendingCompleteId = null;
let installPrompt = null;
const handledReminderAt = new Map();
const localKey = 'daymark-local-tasks-v1';
const localDeviceId = 'daymark-device-id';
let deviceId = localStorage.getItem(localDeviceId);
if (!deviceId) { deviceId = crypto.randomUUID(); localStorage.setItem(localDeviceId, deviceId); }
$('#device-label').textContent = `Device ${deviceId.slice(0, 4).toUpperCase()}`;

function showToast(message) {
  toastNode.textContent = message;
  toastNode.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toastNode.classList.remove('show'), 3200);
}

function intervalMs(task) {
  const n = Number(task.repeatValue);
  return n * ({ minutes: 60_000, hours: 3_600_000, days: 86_400_000 }[task.repeatUnit] || 3_600_000);
}

function localRead() {
  try { return JSON.parse(localStorage.getItem(localKey) || '[]'); } catch { return []; }
}
function localWrite() { localStorage.setItem(localKey, JSON.stringify(tasks)); render(); }
async function createTask(data) {
  tasks.unshift({ ...data, id: crypto.randomUUID(), createdAt: Date.now() });
  localWrite();
}

async function saveTask(id, patch) {
  tasks = tasks.map(t => t.id === id ? { ...t, ...patch } : t);
  localWrite();
}

async function removeTask(id) {
  tasks = tasks.filter(t => t.id !== id);
  localWrite();
}

function dateTime(value) {
  if (!value) return 'Not scheduled';
  const d = new Date(value);
  return d.toLocaleString(undefined, { weekday: 'short', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
}

function escapeHtml(value = '') {
  return String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

function render() {
  const visible = tasks.filter(t => activeFilter === 'completed' ? t.completed : !t.completed)
    .sort((a, b) => (a.nextAt || 0) - (b.nextAt || 0));
  $('#task-count').textContent = tasks.filter(t => !t.completed).length;
  taskList.innerHTML = visible.map(task => {
    const repeatLabel = `Every ${task.repeatValue} ${task.repeatUnit.replace(/s$/, '')}${Number(task.repeatValue) === 1 ? '' : 's'}`;
    return `<article class="task-card ${task.completed ? 'completed-card' : ''}" data-id="${escapeHtml(task.id)}">
      <button class="task-check" data-action="complete" aria-label="Mark ${escapeHtml(task.title)} complete">✓</button>
      <div class="task-main"><h3>${escapeHtml(task.title)}</h3>${task.notes ? `<p class="task-notes">${escapeHtml(task.notes)}</p>` : ''}
        <div class="task-meta"><span>↻ <b>${escapeHtml(repeatLabel)}</b></span><span>♪ <b>${task.soundSeconds}s sound</b></span><span>Started <b>${escapeHtml(dateTime(task.startAt))}</b></span></div>
      </div>
      <div class="task-side"><div><div class="next-label">${task.completed ? 'COMPLETED' : 'NEXT REMINDER'}</div><div class="next-time">${escapeHtml(task.completed ? dateTime(task.completedAt) : dateTime(task.nextAt))}</div></div>
        <div class="task-menu">${task.completed ? '' : `<button class="icon-button" data-action="snooze" title="Remind in 10 minutes">+10 min</button>`}<button class="icon-button" data-action="delete" title="Delete reminder">Delete</button></div>
      </div>
    </article>`;
  }).join('');
  emptyState.classList.toggle('hidden', visible.length > 0);
  taskList.classList.toggle('hidden', visible.length === 0);
}

form.addEventListener('submit', async event => {
  event.preventDefault();
  const startAt = new Date($('#task-start').value).getTime();
  const repeatValue = Number($('#repeat-value').value);
  const soundSeconds = Number($('#sound-value').value);
  if (!Number.isFinite(startAt) || startAt < Date.now() - 60_000) { showToast('Choose a start time in the future.'); return; }
  if (!Number.isInteger(repeatValue) || repeatValue < 1) { showToast('Repeat interval must be at least 1.'); return; }
  if (!Number.isInteger(soundSeconds) || soundSeconds < 1 || soundSeconds > 600) { showToast('Sound duration must be 1–600 seconds.'); return; }
  try {
    await createTask({ title: $('#task-title').value.trim(), notes: $('#task-notes').value.trim(), startAt, nextAt: startAt, repeatValue, repeatUnit: $('#repeat-unit').value, soundSeconds, completed: false, completedAt: null });
    form.reset(); $('#repeat-value').value = '2'; $('#repeat-unit').value = 'hours'; $('#sound-value').value = '20'; setDefaultStart();
    showToast('Reminder saved on this device.');
  } catch (error) { showToast(`Could not save reminder: ${error.message}`); }
});

taskList.addEventListener('click', async event => {
  const button = event.target.closest('[data-action]');
  if (!button) return;
  const card = button.closest('[data-id]');
  const task = tasks.find(t => t.id === card?.dataset.id);
  if (!task) return;
  try {
    if (button.dataset.action === 'complete') {
      pendingCompleteId = task.id; $('#confirm-title').textContent = task.title; $('#confirm-dialog').showModal();
    } else if (button.dataset.action === 'delete') {
      if (confirm(`Delete “${task.title}”?`)) { stopAlarm(task.id); await removeTask(task.id); showToast('Reminder deleted.'); }
    } else if (button.dataset.action === 'snooze') {
      await saveTask(task.id, { nextAt: Date.now() + 10 * 60_000 }); showToast('Reminder moved 10 minutes later.');
    }
  } catch (error) { showToast(`Could not update task: ${error.message}`); }
});

$('#confirm-dialog').addEventListener('close', async () => {
  if ($('#confirm-dialog').returnValue !== 'confirm' || !pendingCompleteId) return;
  try {
    await saveTask(pendingCompleteId, { completed: true, completedAt: Date.now() });
    stopAlarm(pendingCompleteId); showToast('Done. Its reminders have stopped.');
  } catch (error) { showToast(`Could not mark complete: ${error.message}`); }
  pendingCompleteId = null;
});

document.querySelectorAll('.filter').forEach(button => button.addEventListener('click', () => {
  activeFilter = button.dataset.filter;
  document.querySelectorAll('.filter').forEach(b => b.classList.toggle('active', b === button)); render();
}));

function startAlarm(task) {
  const now = Date.now();
  const endAt = now + Number(task.soundSeconds) * 1000;
  if (!audioContext) audioContext = new (window.AudioContext || window.webkitAudioContext)();
  audioContext.resume().catch(() => {});
  stopAlarm(task.id);
  if (Notification.permission === 'granted' && document.hidden) new Notification('Daymark reminder', { body: task.title, tag: `daymark-${task.id}`, renotify: true });
  const timer = setInterval(() => {
    if (Date.now() >= endAt) { stopAlarm(task.id); return; }
    try {
      const osc = audioContext.createOscillator(); const gain = audioContext.createGain();
      osc.type = 'sine'; osc.frequency.value = 740; gain.gain.value = 0.045;
      osc.connect(gain); gain.connect(audioContext.destination); osc.start(); osc.stop(audioContext.currentTime + 0.16);
    } catch { /* audio may be blocked until the user interacts with the app */ }
  }, 900);
  alarmTimers.set(task.id, timer);
  if (!document.hidden) showToast(`Reminder: ${task.title}`);
}

function stopAlarm(taskId) { const timer = alarmTimers.get(taskId); if (timer) clearInterval(timer); alarmTimers.delete(taskId); }
function scheduleCheck() {
  const now = Date.now();
  for (const task of tasks) {
    if (task.completed || !task.nextAt || task.nextAt > now) continue;
    if (handledReminderAt.get(task.id) === task.nextAt) continue;
    handledReminderAt.set(task.id, task.nextAt);
    const ms = intervalMs(task);
    const missed = Math.floor((now - task.nextAt) / ms) + 1;
    const nextAt = task.nextAt + missed * ms;
    saveTask(task.id, { nextAt }).catch(error => showToast(`Could not update reminder: ${error.message}`));
    startAlarm(task);
  }
}
setInterval(scheduleCheck, 12_000);

$('#enable-notifications').addEventListener('click', async () => {
  try {
    if (!audioContext) audioContext = new (window.AudioContext || window.webkitAudioContext)();
    await audioContext.resume();
  } catch { /* Browser may not provide Web Audio. */ }
  if (!('Notification' in window)) { showToast('This browser does not support notifications.'); return; }
  const permission = await Notification.requestPermission();
  showToast(permission === 'granted' ? 'Browser alerts enabled.' : 'Allow notifications in browser settings to receive alerts.');
});

window.addEventListener('beforeinstallprompt', event => { event.preventDefault(); installPrompt = event; $('#install-app').classList.remove('hidden'); });
$('#install-app').addEventListener('click', async () => { if (!installPrompt) return; await installPrompt.prompt(); installPrompt = null; $('#install-app').classList.add('hidden'); });
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('./sw.js').catch(() => {});

function setDefaultStart() {
  const d = new Date(Date.now() + 60_000); d.setSeconds(0, 0);
  const local = new Date(d.getTime() - d.getTimezoneOffset() * 60_000).toISOString().slice(0, 16);
  $('#task-start').value = local;
}
$('#today-label').textContent = new Date().toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' });
setDefaultStart();
tasks = localRead();
render();
