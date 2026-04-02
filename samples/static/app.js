/**
 * office-connect sample — Vue 3 + Quasar (Outlook-style)
 *
 * Third-party libraries (bundled in vendor/):
 *   Vue.js v3.5.30    — MIT License    — https://github.com/vuejs/core
 *   Quasar v2.18.6    — MIT License    — https://github.com/quasarframework/quasar
 *   Phosphor Icons     — MIT License    — https://github.com/phosphor-icons/core
 *   Material Icons     — Apache 2.0     — https://github.com/google/material-design-icons
 */

const { createApp, ref, reactive, computed, onMounted, watch, nextTick } = Vue;

/* ── Avatar colour palette ───────────────────────────────────── */
const AVATAR_COLORS = [
  '#0078d4','#008272','#5c2d91','#c239b3','#e3008c',
  '#986f0b','#498205','#004e8c','#8764b8','#ca5010',
  '#7a7574','#004b50','#69797e','#a4262c','#0063b1',
];

function hashStr(s) {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  return Math.abs(h);
}

/* ── Folder icon mapping ─────────────────────────────────────── */
const FOLDER_ICONS = {
  inbox: 'ph ph-tray',
  drafts: 'ph ph-pencil-simple',
  'sent items': 'ph ph-paper-plane-tilt',
  'deleted items': 'ph ph-trash',
  archive: 'ph ph-archive-box',
  'junk email': 'ph ph-warning-circle',
  spam: 'ph ph-warning-circle',
  outbox: 'ph ph-upload-simple',
  notes: 'ph ph-note',
  'conversation history': 'ph ph-chat-dots',
  rss: 'ph ph-rss',
};

const MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];

const app = createApp({
  setup() {
    // ── core state ─────────────────────────────────────────
    const authenticated = ref(false);
    const csrfToken = ref('');
    const view = ref('mail');

    // ── mail state ─────────────────────────────────────────
    const folders = ref([]);
    const selectedFolder = ref('');
    const currentFolderName = ref('Inbox');
    const messages = ref([]);
    const filteredMessages = ref([]);
    const mailSearch = ref('');
    const selectedMessage = ref(null);
    const loadingMail = ref(false);

    // ── compose state ──────────────────────────────────────
    const composeMode = ref(false);
    const sending = ref(false);
    const compose = reactive({
      toList: [],
      toInput: '',
      subject: '',
      body: '',
      replyTo: null,
      isForward: false,
    });
    const toSuggestions = ref([]);
    let toSearchTimer = null;

    // ── calendar state ─────────────────────────────────────
    const now = new Date();
    const calYear = ref(now.getFullYear());
    const calMonth = ref(now.getMonth());
    const miniCalYear = ref(now.getFullYear());
    const miniCalMonth = ref(now.getMonth());
    const events = ref([]);
    const loadingCal = ref(false);
    const selectedCalDate = ref('');

    // ── people state ───────────────────────────────────────
    const people = ref([]);
    const peopleQuery = ref('');
    const loadingPeople = ref(false);
    const failedPhotos = ref(new Set());
    let peopleTimer = null;

    // ── profile ────────────────────────────────────────────
    const profile = ref(null);

    // ════════════════════════════════════════════════════════
    // HELPERS
    // ════════════════════════════════════════════════════════

    async function api(path, opts) {
      const headers = { 'Content-Type': 'application/json' };
      if (csrfToken.value) headers['X-CSRF-Token'] = csrfToken.value;
      const res = await fetch(path, { ...opts, headers: { ...headers, ...(opts || {}).headers } });
      if (res.status === 401) { authenticated.value = false; throw new Error('Not authenticated'); }
      if (!res.ok) { const t = await res.text(); throw new Error(t || res.statusText); }
      return res.json();
    }

    function toast(msg, type) {
      Quasar.Notify.create({
        message: msg,
        color: type === 'error' ? 'negative' : type === 'warning' ? 'warning' : 'positive',
        position: 'bottom-right', timeout: 3000,
      });
    }

    function avatarColor(name) { return AVATAR_COLORS[hashStr(name || '?') % AVATAR_COLORS.length]; }

    function initials(name) {
      if (!name) return '?';
      const parts = name.trim().split(/[\s@.]+/);
      if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
      return name.slice(0, 2).toUpperCase();
    }

    function folderIcon(name) {
      return FOLDER_ICONS[(name || '').toLowerCase()] || 'ph ph-folder';
    }

    function formatShortDate(iso) {
      if (!iso) return '';
      const d = new Date(iso);
      const today = new Date();
      if (d.toDateString() === today.toDateString()) {
        return d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
      }
      const yesterday = new Date(today); yesterday.setDate(today.getDate() - 1);
      if (d.toDateString() === yesterday.toDateString()) return 'Yesterday';
      return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
    }

    function formatFullDate(iso) {
      if (!iso) return '';
      return new Date(iso).toLocaleString(undefined, {
        weekday: 'long', year: 'numeric', month: 'long', day: 'numeric',
        hour: '2-digit', minute: '2-digit',
      });
    }

    function formatTime(iso) {
      if (!iso) return '';
      return new Date(iso).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
    }

    function filterMessages() {
      const q = mailSearch.value.toLowerCase().trim();
      if (!q) { filteredMessages.value = messages.value; return; }
      filteredMessages.value = messages.value.filter(m =>
        (m.subject || '').toLowerCase().includes(q) ||
        (m.from_name || '').toLowerCase().includes(q) ||
        (m.from_email || '').toLowerCase().includes(q) ||
        (m.preview || '').toLowerCase().includes(q)
      );
    }

    // ════════════════════════════════════════════════════════
    // AUTH
    // ════════════════════════════════════════════════════════

    async function checkAuth() {
      try {
        const data = await api('auth-status');
        authenticated.value = data.authenticated;
        if (data.authenticated) {
          const csrf = await api('csrf-token');
          csrfToken.value = csrf.token;
          loadProfile();
          loadFolders();
        }
      } catch { /* not logged in */ }
    }

    async function loadProfile() {
      try { profile.value = await api('api/profile'); } catch { /* ignore */ }
    }

    // ════════════════════════════════════════════════════════
    // MAIL
    // ════════════════════════════════════════════════════════

    async function loadFolders() {
      try {
        folders.value = await api('api/mail/folders');
        const inbox = folders.value.find(f => f.name.toLowerCase() === 'inbox');
        if (inbox) { selectedFolder.value = inbox.id; currentFolderName.value = inbox.name; }
        else if (folders.value.length) { selectedFolder.value = folders.value[0].id; currentFolderName.value = folders.value[0].name; }
        loadMessages();
      } catch (e) { toast('Failed to load folders: ' + e.message, 'error'); }
    }

    async function selectFolder(id) {
      selectedFolder.value = id;
      const f = folders.value.find(x => x.id === id);
      currentFolderName.value = f ? f.name : '';
      selectedMessage.value = null;
      await loadMessages();
    }

    async function loadMessages() {
      loadingMail.value = true;
      try {
        const data = await api('api/mail/messages?folder_id=' + encodeURIComponent(selectedFolder.value) + '&limit=50');
        messages.value = data.messages;
        filterMessages();
      } catch (e) { toast('Failed to load messages: ' + e.message, 'error'); }
      finally { loadingMail.value = false; }
    }

    async function openMessage(id) {
      try {
        selectedMessage.value = await api('api/mail/messages/' + encodeURIComponent(id));
      } catch (e) { toast('Failed to load message: ' + e.message, 'error'); }
    }

    // ════════════════════════════════════════════════════════
    // COMPOSE
    // ════════════════════════════════════════════════════════

    function startCompose(prefillTo) {
      compose.toList = prefillTo ? [prefillTo] : [];
      compose.toInput = '';
      compose.subject = '';
      compose.body = '';
      compose.replyTo = null;
      compose.isForward = false;
      toSuggestions.value = [];
      composeMode.value = true;
    }

    function startReply(replyAll) {
      if (!selectedMessage.value) return;
      const m = selectedMessage.value;
      compose.toList = [m.from_email];
      compose.toInput = '';
      compose.subject = 'Re: ' + (m.subject || '');
      compose.body = '';
      compose.replyTo = m.email_id;
      compose.isForward = false;
      toSuggestions.value = [];
      composeMode.value = true;
    }

    function startForward() {
      if (!selectedMessage.value) return;
      const m = selectedMessage.value;
      compose.toList = [];
      compose.toInput = '';
      compose.subject = 'Fw: ' + (m.subject || '');
      compose.body = '\n\n--- Forwarded message ---\n' + (m.body_preview || m.body || '');
      compose.replyTo = null;
      compose.isForward = true;
      toSuggestions.value = [];
      composeMode.value = true;
    }

    function addRecipient() {
      const val = compose.toInput.trim();
      if (val && !compose.toList.includes(val)) compose.toList.push(val);
      compose.toInput = '';
      toSuggestions.value = [];
    }

    function pickSuggestion(p) {
      if (p.email && !compose.toList.includes(p.email)) compose.toList.push(p.email);
      compose.toInput = '';
      toSuggestions.value = [];
    }

    async function onToInput() {
      clearTimeout(toSearchTimer);
      const q = compose.toInput.trim();
      if (q.length < 2) { toSuggestions.value = []; return; }
      toSearchTimer = setTimeout(async () => {
        try {
          toSuggestions.value = await api('api/people/search?q=' + encodeURIComponent(q));
        } catch { toSuggestions.value = []; }
      }, 300);
    }

    async function sendCompose() {
      if (!compose.toList.length || !compose.subject) {
        toast('To and Subject are required', 'warning'); return;
      }
      sending.value = true;
      try {
        if (compose.replyTo && !compose.isForward) {
          await api('api/mail/reply', {
            method: 'POST',
            body: JSON.stringify({ message_id: compose.replyTo, body: compose.body, reply_all: false }),
          });
        } else {
          await api('api/mail/send', {
            method: 'POST',
            body: JSON.stringify({ to: compose.toList, subject: compose.subject, body: compose.body }),
          });
        }
        toast('Message sent');
        composeMode.value = false;
        loadMessages();
      } catch (e) { toast('Send failed: ' + e.message, 'error'); }
      finally { sending.value = false; }
    }

    function discardCompose() { composeMode.value = false; }

    // ════════════════════════════════════════════════════════
    // CALENDAR
    // ════════════════════════════════════════════════════════

    const calMonthName = computed(() => MONTHS[calMonth.value]);
    const miniCalMonthName = computed(() => MONTHS[miniCalMonth.value]);

    function isSameDay(a, b) {
      return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
    }

    function buildMonthGrid(year, month) {
      const first = new Date(year, month, 1);
      const last = new Date(year, month + 1, 0);
      const start = new Date(first);
      start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
      const today = new Date();
      const weeks = [];
      const cur = new Date(start);
      for (let w = 0; w < 6; w++) {
        const week = [];
        for (let d = 0; d < 7; d++) {
          const dt = new Date(cur);
          const dayEvents = events.value.filter(e => {
            const eStart = new Date(e.start_time);
            const eEnd = new Date(e.end_time);
            return dt >= new Date(eStart.getFullYear(), eStart.getMonth(), eStart.getDate()) &&
                   dt <= new Date(eEnd.getFullYear(), eEnd.getMonth(), eEnd.getDate());
          });
          week.push({
            date: dt.toISOString().slice(0, 10),
            day: dt.getDate(),
            currentMonth: dt.getMonth() === month,
            isToday: isSameDay(dt, today),
            events: dayEvents,
          });
          cur.setDate(cur.getDate() + 1);
        }
        weeks.push(week);
      }
      return weeks;
    }

    const calendarGrid = computed(() => buildMonthGrid(calYear.value, calMonth.value));

    function buildMiniCal(year, month) {
      const first = new Date(year, month, 1);
      const start = new Date(first);
      start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
      const today = new Date();
      const days = [];
      const cur = new Date(start);
      for (let i = 0; i < 42; i++) {
        days.push({
          date: cur.toISOString().slice(0, 10),
          day: cur.getDate(),
          currentMonth: cur.getMonth() === month,
          isToday: isSameDay(cur, today),
        });
        cur.setDate(cur.getDate() + 1);
      }
      return days;
    }

    const miniCalDays = computed(() => buildMiniCal(miniCalYear.value, miniCalMonth.value));

    function calPrev() {
      calMonth.value--;
      if (calMonth.value < 0) { calMonth.value = 11; calYear.value--; }
      miniCalMonth.value = calMonth.value;
      miniCalYear.value = calYear.value;
      loadEvents();
    }
    function calNext() {
      calMonth.value++;
      if (calMonth.value > 11) { calMonth.value = 0; calYear.value++; }
      miniCalMonth.value = calMonth.value;
      miniCalYear.value = calYear.value;
      loadEvents();
    }
    function miniCalPrev() {
      miniCalMonth.value--;
      if (miniCalMonth.value < 0) { miniCalMonth.value = 11; miniCalYear.value--; }
    }
    function miniCalNext() {
      miniCalMonth.value++;
      if (miniCalMonth.value > 11) { miniCalMonth.value = 0; miniCalYear.value++; }
    }
    function goToToday() {
      const t = new Date();
      calYear.value = t.getFullYear(); calMonth.value = t.getMonth();
      miniCalYear.value = t.getFullYear(); miniCalMonth.value = t.getMonth();
      selectedCalDate.value = t.toISOString().slice(0, 10);
      loadEvents();
    }
    function goToDate(dateStr) {
      const d = new Date(dateStr);
      calYear.value = d.getFullYear(); calMonth.value = d.getMonth();
      selectedCalDate.value = dateStr;
      loadEvents();
    }

    async function loadEvents() {
      loadingCal.value = true;
      try {
        const start = new Date(calYear.value, calMonth.value, 1).toISOString().slice(0, 10);
        const end = new Date(calYear.value, calMonth.value + 1, 6).toISOString().slice(0, 10);
        const data = await api('api/calendar/events?start=' + start + '&end=' + end);
        events.value = data.events || [];
      } catch (e) { toast('Failed to load events: ' + e.message, 'error'); }
      finally { loadingCal.value = false; }
    }

    // auto-load calendar when switching to calendar view
    watch(view, (v) => {
      if (v === 'calendar' && events.value.length === 0) loadEvents();
    });

    // ════════════════════════════════════════════════════════
    // PEOPLE
    // ════════════════════════════════════════════════════════

    async function loadPeople(q) {
      loadingPeople.value = true;
      try {
        people.value = await api('api/people/search?q=' + encodeURIComponent(q || ''));
      } catch (e) { toast('Search failed: ' + e.message, 'error'); }
      finally { loadingPeople.value = false; }
    }

    function debouncePeopleSearch() {
      clearTimeout(peopleTimer);
      peopleTimer = setTimeout(() => loadPeople(peopleQuery.value), 400);
    }

    function personPhotoUrl(p) {
      if (failedPhotos.value.has(p.id)) return null;
      return 'api/people/' + encodeURIComponent(p.id) + '/photo';
    }

    function onPhotoError(p) {
      failedPhotos.value.add(p.id);
    }

    function startMailTo(p) {
      view.value = 'mail';
      startCompose(p.email);
    }

    // ════════════════════════════════════════════════════════
    // INIT
    // ════════════════════════════════════════════════════════

    onMounted(() => { checkAuth(); });

    return {
      // core
      authenticated, profile, view,
      // mail
      folders, selectedFolder, currentFolderName, messages, filteredMessages,
      mailSearch, selectedMessage, loadingMail,
      selectFolder, openMessage, filterMessages,
      // compose
      composeMode, sending, compose, toSuggestions,
      startCompose, startReply, startForward,
      addRecipient, pickSuggestion, onToInput,
      sendCompose, discardCompose,
      // calendar
      calYear, calMonth, calMonthName, miniCalYear, miniCalMonth, miniCalMonthName,
      events, loadingCal, selectedCalDate, calendarGrid, miniCalDays,
      calPrev, calNext, miniCalPrev, miniCalNext, goToToday, goToDate, loadEvents,
      // people
      people, peopleQuery, loadingPeople,
      loadPeople, debouncePeopleSearch, personPhotoUrl, onPhotoError, startMailTo,
      // helpers
      avatarColor, initials, folderIcon,
      formatShortDate, formatFullDate, formatTime,
    };
  },
});

app.use(Quasar, {
  config: { notify: { position: 'bottom-right', timeout: 3000 } },
});
app.mount('#q-app');
