/**
 * Lara Smile · Event Takip API
 * -------------------------------------------------------------
 * Bu script randevu sheet'ini okur ve web sitesine SADECE sayıları
 * (event / gün / sorumlu bazında randevu-geldi-book) JSON olarak verir.
 * Hasta adı, lead numarası, tedavi notu dışarı ÇIKMAZ.
 *
 * Kurulum: KURULUM.md dosyasına bak.
 */

// Ek dosyalar: başka bir Google Sheet'teki event'leri de dahil etmek
// istersen ID'sini buraya ekle. Bu dosyanın kendisi her zaman okunur.
// Örnek: const EXTRA_SPREADSHEET_IDS = ['1AbC...xyz'];
const EXTRA_SPREADSHEET_IDS = [];

// Sorumlu adı düzeltmeleri (sheet'te farklı yazılırsa tek isimde toplar)
// Anahtar küçük harf yazılır.
const AGENT_ALIASES = {
  // 'cemre |': 'Cemre',
};

const CACHE_SECONDS = 30;

// Aynı event'te aynı isim iki kez yazılmışsa tek say? (false = her satır ayrı randevu)
const MERGE_DUPLICATES = false;

/* ============================================================= */

function doGet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('events_v1');
  let json = hit;
  if (!hit) {
    json = JSON.stringify(buildPayload_());
    try { cache.put('events_v1', json, CACHE_SECONDS); } catch (e) {}
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function buildPayload_() {
  const books = [SpreadsheetApp.getActiveSpreadsheet()];
  EXTRA_SPREADSHEET_IDS.forEach(function (id) {
    try { books.push(SpreadsheetApp.openById(id)); } catch (e) {}
  });

  const events = [];
  books.forEach(function (ss) {
    const tz = ss.getSpreadsheetTimeZone();
    ss.getSheets().forEach(function (sh) {
      const name = sh.getName();
      if (name.charAt(0) === '_') return;          // "_" ile başlayan sekmeler atlanır
      if (sh.isSheetHidden()) return;
      const values = sh.getDataRange().getValues();
      const ev = parseEventSheet(values, name, function (d) {
        return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      });
      if (ev) events.push(ev);
    });
  });

  events.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : a.type < b.type ? -1 : 1; });
  return { updatedAt: new Date().toISOString(), events: events };
}

/* ---------- Ayrıştırıcı (sheet düzeninden bağımsız) ----------
 * Her gün bloğu "Book Saati" başlığıyla başlar.
 * Blok içindeki sütunlar başlık adına göre bulunur:
 *   Book Saati | Lead Numarası | Hasta Adı | Sorumlu | İlgilendiği Tedavi | Durum
 * 2. satırda blok tarihi var. Boş başlıklı sütun bloğu bitirir.
 */
function parseEventSheet(values, sheetName, fmtDate, log) {
  log = log || null;
  if (!values || values.length < 3) return null;
  const head = values[0].map(norm_);
  const starts = [];
  head.forEach(function (h, i) { if (h.indexOf('saat') !== -1) starts.push(i); });
  if (!starts.length) return null;

  const days = [];
  const seen = {};           // aynı event'te aynı hasta iki kez yazılmışsa tek say
  let duplicates = 0;
  const RANK = { '': 0, gelmedi: 1, geldi: 2, book: 3 };

  const note = function (msg) { if (log) log.push(sheetName + ' | ' + msg); };

  starts.forEach(function (s, bi) {
    const end = bi + 1 < starts.length ? starts[bi + 1] : head.length;
    const col = {};
    for (let c = s + 1; c < end; c++) {
      const h = head[c];
      if (!h) break;
      if (h.indexOf('lead') !== -1) col.lead = c;
      else if (h.indexOf('hasta') !== -1 || h.indexOf('isim') !== -1) col.name = c;
      else if (h.indexOf('sorumlu') !== -1) col.agent = c;
      else if (h.indexOf('tedavi') !== -1) col.treat = c;
      else if (h.indexOf('durum') !== -1) col.status = c;
    }
    if (col.name === undefined) { note('Sütun ' + colName_(s) + ': "Hasta Adı" başlığı bulunamadı, blok atlandı'); return; }

    const dateVal = toDate_(values[1][s]);
    if (!dateVal) { note('Sütun ' + colName_(s) + ': 2. satırda tarih yok, blok atlandı'); return; }
    const day = { date: fmtDate(dateVal), slots: 0, randevu: 0, geldi: 0, book: 0, gelmedi: 0, agents: {} };

    for (let r = 2; r < values.length; r++) {
      const row = values[r];
      const hasTime = row[s] !== '' && row[s] !== null;
      let patient = clean_(row[col.name]);
      const lead = col.lead !== undefined ? clean_(row[col.lead]) : '';
      if (hasTime) day.slots++;
      if (!patient && lead) { patient = 'lead ' + lead; note(day.date + ' satır ' + (r + 1) + ': isim boş, lead no ile sayıldı (' + lead + ')'); }
      if (!patient) continue;

      const status = col.status !== undefined ? statusOf_(row[col.status]) : '';
      if (status === 'iptal') { note(day.date + ' satır ' + (r + 1) + ': İptal, sayılmadı (' + patient + ')'); continue; }

      const agent = agentName_(col.agent !== undefined ? row[col.agent] : '');
      const key = norm_(patient);
      if (seen[key]) note(day.date + ' satır ' + (r + 1) + ': tekrar eden isim (' + patient + ')' + (MERGE_DUPLICATES ? ', tek sayıldı' : ', ayrı sayıldı'));
      if (seen[key] && MERGE_DUPLICATES) {           // tekrar eden kayıt
        duplicates++;
        const prev = seen[key];
        if (RANK[status] > RANK[prev.status]) {
          bump_(prev.day, prev.agent, prev.status, -1);
          bump_(prev.day, prev.agent, status, +1);
          prev.status = status;
        }
        continue;
      }
      seen[key] = { day: day, agent: agent, status: status };
      day.randevu++;
      day.agents[agent] = day.agents[agent] || { randevu: 0, geldi: 0, book: 0, gelmedi: 0 };
      day.agents[agent].randevu++;
      bump_(day, agent, status, +1);
    }
    days.push(day);
  });

  if (!days.length) return null;
  days.sort(function (a, b) { return a.date < b.date ? -1 : 1; });

  const ct = cityAndType_(sheetName);
  const hasStatus = head.some(function (h) { return h.indexOf('durum') !== -1; });

  return {
    id: slug_(sheetName + '-' + days[0].date),
    name: sheetName,
    city: ct.city,
    type: ct.type,
    start: days[0].date,
    end: days[days.length - 1].date,
    month: days[0].date.slice(0, 7),
    hasStatusColumn: hasStatus,
    duplicates: duplicates,
    days: days
  };
}

/* Sekme adından şehir ve tür:
 *   "Manchester"              -> Manchester / Dental
 *   "Manchester Estetik"      -> Manchester / Estetik
 *   "Manchester Physiotherapy"-> Manchester / Physiotherapy
 *   "Milton Keynes - Implant" -> Milton Keynes / Implant   (tire ile istediğin türü yazabilirsin)
 */
const TYPE_WORDS = ['estetik', 'aesthetic', 'aesthetics', 'physiotherapy', 'physio', 'fizyoterapi', 'fizyo', 'dental'];
function cityAndType_(name) {
  const n = clean_(name);
  const dash = n.split(/\s+[-–]\s+/);
  if (dash.length > 1) return { city: dash[0], type: dash.slice(1).join(' ') };
  const words = n.split(' ');
  const typeW = [], cityW = [];
  words.forEach(function (w) { (TYPE_WORDS.indexOf(w.toLocaleLowerCase('tr')) !== -1 ? typeW : cityW).push(w); });
  const type = typeW.filter(function (w) { return w.toLocaleLowerCase('tr') !== 'dental'; }).join(' ');
  return {
    city: cityW.join(' ') || n,
    type: type ? type.charAt(0).toLocaleUpperCase('tr') + type.slice(1) : 'Dental'
  };
}

function bump_(day, agent, status, d) {
  if (!status) return;
  const a = day.agents[agent];
  if (status === 'book') { day.book += d; a.book += d; day.geldi += d; a.geldi += d; }
  else if (status === 'geldi') { day.geldi += d; a.geldi += d; }
  else if (status === 'gelmedi') { day.gelmedi += d; a.gelmedi += d; }
}

function statusOf_(v) {
  const s = norm_(v);
  if (!s) return '';
  if (s.indexOf('iptal') !== -1 || s.indexOf('cancel') !== -1) return 'iptal';
  if (s.indexOf('gelmedi') !== -1 || s.indexOf('no show') !== -1 || s.indexOf('noshow') !== -1) return 'gelmedi';
  if (s.indexOf('book') !== -1 || s.indexOf('depozit') !== -1 || s.indexOf('satış') !== -1 || s.indexOf('satis') !== -1) return 'book';
  if (s.indexOf('geldi') !== -1 || s.indexOf('show') !== -1) return 'geldi';
  return '';
}

function agentName_(v) {
  let s = clean_(v).replace(/[|/\\]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!s) return 'Atanmamış';
  const alias = AGENT_ALIASES[norm_(s)];
  if (alias) return alias;
  s = s.split(' ')[0];                               // "Cemre |" -> "Cemre"
  s = s.replace(/İ/g, 'i').replace(/I/g, 'i').toLowerCase();
  return s.charAt(0).toLocaleUpperCase('tr') + s.slice(1);
}

function toDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
    if (m) return new Date(+(m[3].length === 2 ? '20' + m[3] : m[3]), +m[2] - 1, +m[1]);
  }
  return null;
}

function colName_(i) { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
function clean_(v) { return v === null || v === undefined ? '' : String(v).replace(/\s+/g, ' ').trim(); }
function norm_(v) { return clean_(v).toLocaleLowerCase('tr'); }
function slug_(s) {
  return norm_(s).replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u').replace(/ş/g, 's')
    .replace(/ö/g, 'o').replace(/ç/g, 'c').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

/* Sayılar sheet'le tutmuyorsa: Çalıştır > teshis, sonra Günlükler.
 * Her sekme için sayılan randevu ve atlanan/farklı sayılan satırları listeler.
 * (Sadece editörde görünür, siteye gitmez.) */
function teshis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tz = ss.getSpreadsheetTimeZone();
  ss.getSheets().forEach(function (sh) {
    const log = [];
    const ev = parseEventSheet(sh.getDataRange().getValues(), sh.getName(),
      function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }, log);
    if (!ev) { Logger.log(sh.getName() + ': event olarak okunmadı'); return; }
    const total = ev.days.reduce(function (a, d) { return a + d.randevu; }, 0);
    Logger.log('== ' + sh.getName() + ': ' + total + ' randevu (' + ev.days.map(function (d) { return d.date + ': ' + d.randevu; }).join(', ') + ')');
    log.forEach(function (l) { Logger.log('   ' + l); });
  });
}

/* Apps Script editöründe test için: Çalıştır > testPayload, sonra Günlükler */
function testPayload() {
  Logger.log(JSON.stringify(buildPayload_(), null, 2));
}
