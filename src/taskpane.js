/* WordFormatSearch – לוגיקת התוסף
 * חיפוש טקסט / עיצוב / טקסט+עיצוב, ובחירת Range אמיתי ב-Word.
 * התוסף אינו משנה את תוכן המסמך או את עיצובו – רק קורא ובוחר.
 */
"use strict";

/* =========================================================
 * 1. קבועים
 * ========================================================= */

// מפרידי מילים עבור getTextRanges().
// גרש/גרשיים/מרכאות אינם מפרידים – כך "עו״ד" נשאר מילה אחת;
// בקצוות המילה הם מוסרים ע"י extractCore().
const WORD_DELIMITERS = [
  " ", "\t", "\n", "\r", "\u00A0", "\u000B",
  ".", ",", ";", ":", "!", "?",
  "(", ")", "[", "]", "{", "}",
  "-", "–", "—", "־", "/", "\\", "…"
];

// תווים שמוסרים מתחילת/סוף מילה
const EDGE_CHARS = "\\s\u00A0.,;:!?()\\[\\]{}\"'״׳`“”„‘’\\-–—־/\\\\…";
const EDGE_RE = new RegExp("^[" + EDGE_CHARS + "]+|[" + EDGE_CHARS + "]+$", "g");

const PARAGRAPH_BATCH = 150;   // פסקאות בכל sync בחיפוש עיצוב
const MAX_RESULTS = 10000;     // הגבלת בטיחות
const MAX_SEARCH_LEN = 255;    // מגבלת Word Search

// צבעי הדגשה של Word → Hex
const HIGHLIGHT_HEX = {
  yellow: "#FFFF00", brightgreen: "#00FF00", lime: "#00FF00",
  turquoise: "#00FFFF", cyan: "#00FFFF", pink: "#FF00FF", magenta: "#FF00FF",
  blue: "#0000FF", red: "#FF0000", darkblue: "#000080", navy: "#000080",
  teal: "#008080", green: "#008000", violet: "#800080", purple: "#800080",
  darkred: "#800000", maroon: "#800000", darkyellow: "#808000", olive: "#808000",
  gray50: "#808080", darkgray: "#808080", gray: "#808080",
  gray25: "#C0C0C0", lightgray: "#C0C0C0", silver: "#C0C0C0",
  black: "#000000", white: "#FFFFFF"
};

const HIGHLIGHT_OPTIONS = [
  ["__any__", "הדגשה כלשהי"],
  ["none", "ללא הדגשה"],
  ["#FFFF00", "צהוב"],
  ["#00FF00", "ירוק בהיר"],
  ["#00FFFF", "טורקיז"],
  ["#FF00FF", "ורוד"],
  ["#0000FF", "כחול"],
  ["#FF0000", "אדום"],
  ["#000080", "כחול כהה"],
  ["#008080", "כחלחל"],
  ["#008000", "ירוק"],
  ["#800080", "סגול"],
  ["#800000", "אדום כהה"],
  ["#808000", "צהוב כהה"],
  ["#808080", "אפור 50%"],
  ["#C0C0C0", "אפור 25%"],
  ["#000000", "שחור"]
];

const UNDERLINE_OPTIONS = [
  ["__any__", "קו תחתון כלשהו"],
  ["None", "ללא קו תחתון"],
  ["Single", "קו יחיד"],
  ["Double", "קו כפול"],
  ["Thick", "קו עבה"],
  ["Dotted", "נקודות"],
  ["DashLine", "מקווקו"],
  ["Wave", "גלי"],
  ["Word", "מילים בלבד"]
];

const PROPS = ["bold", "italic", "underline", "name", "size", "color", "highlightColor"];

/* =========================================================
 * 2. פונקציות טהורות (נבדקות גם ב-Node)
 * ========================================================= */

function normalizeHex(value) {
  if (value === null || value === undefined) return null;
  let s = String(value).trim();
  if (!s) return null;
  if (s[0] !== "#") s = "#" + s;
  if (/^#[0-9a-f]{3}$/i.test(s)) s = "#" + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
  return /^#[0-9a-f]{6}$/i.test(s) ? s.toUpperCase() : null;
}

/** מחזיר "none", ערך Hex, או null אם לא ידוע */
function normalizeHighlight(value) {
  if (value === null || value === undefined) return "none";
  const s = String(value).trim();
  if (!s) return "none";
  const low = s.toLowerCase().replace(/[\s_]/g, "");
  if (low === "none" || low === "nohighlight" || low === "transparent" || low === "auto") return "none";
  if (HIGHLIGHT_HEX[low]) return HIGHLIGHT_HEX[low];
  return normalizeHex(s);
}

function normalizeFontName(value) {
  return value ? String(value).trim().toLowerCase() : "";
}

/** האם ערך שנקרא מ-Word מייצג עיצוב מעורב */
function isMixed(prop, value) {
  if (prop === "highlightColor") return false; // null = ללא הדגשה
  if (value === null || value === undefined) return true;
  if (prop === "underline" && String(value).toLowerCase() === "mixed") return true;
  if ((prop === "name" || prop === "color") && value === "") return true;
  return false;
}

/** criteria: { bold?:bool, italic?:bool, underline?:string, name?:string,
 *              size?:number, color?:string, highlightColor?:string }
 *  רק מפתחות קיימים משתתפים בהשוואה. */
function formattingMatches(font, criteria) {
  for (const prop of Object.keys(criteria)) {
    const want = criteria[prop];
    const have = font[prop];
    if (isMixed(prop, have)) return false;
    switch (prop) {
      case "bold":
      case "italic":
        if (Boolean(have) !== want) return false;
        break;
      case "underline":
        if (want === "__any__") { if (String(have) === "None") return false; }
        else if (String(have).toLowerCase() !== String(want).toLowerCase()) return false;
        break;
      case "name":
        if (normalizeFontName(have) !== normalizeFontName(want)) return false;
        break;
      case "size":
        if (Math.abs(Number(have) - Number(want)) > 0.05) return false;
        break;
      case "color":
        if (normalizeHex(have) !== normalizeHex(want)) return false;
        break;
      case "highlightColor": {
        const h = normalizeHighlight(have);
        if (want === "__any__") { if (h === "none") return false; }
        else if (h !== want) return false;
        break;
      }
    }
  }
  return true;
}

/** האם יש בתוצאה ערך מעורב באחד המאפיינים הנבדקים */
function hasMixedCriteria(font, criteria) {
  return Object.keys(criteria).some((p) => isMixed(p, font[p]));
}

/** מסיר סימני פיסוק ורווחים מקצוות המילה */
function extractCore(text) {
  return String(text || "").replace(EDGE_RE, "");
}

function hasAnyFormattingCriteria(criteria) {
  return Object.keys(criteria).length > 0;
}

/** Word Search מפרש ^ כקוד מיוחד */
function escapeWordSearch(text) {
  return String(text).replace(/\^/g, "^^");
}

function buildFontLoadString(criteria, prefix) {
  const keys = Object.keys(criteria);
  const parts = [prefix + "text"];
  keys.forEach((k) => parts.push(prefix + "font/" + k));
  return parts.join(",");
}

/**
 * מאחד מילים תואמות רצופות באותה פסקה לתוצאה אחת (כמו Find של Word לפי עיצוב).
 * matches: [{ para, tokenIndex, core, raw, breakBefore, breakAfter }] לפי סדר המסמך.
 * רצף נשבר כשיש מילה לא תואמת ביניהן, בסוף פסקה, או בסימן פיסוק בעיצוב שונה.
 */
function groupRuns(matches) {
  const runs = [];
  let cur = null;
  for (const m of matches) {
    const last = cur && cur.items[cur.items.length - 1];
    if (cur && m.para === cur.para && m.tokenIndex === last.tokenIndex + 1 && !last.breakAfter && !m.breakBefore) {
      cur.items.push(m);
    } else {
      cur = { para: m.para, items: [m] };
      runs.push(cur);
    }
  }
  return runs;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    normalizeHex, normalizeHighlight, isMixed, formattingMatches,
    hasMixedCriteria, extractCore, escapeWordSearch, hasAnyFormattingCriteria,
    buildFontLoadString, groupRuns, WORD_DELIMITERS
  };
}

/* =========================================================
 * 3. State
 * ========================================================= */

const state = {
  results: [],        // [{ text, locator, para }]
  currentIndex: -1,
  searchText: "",
  matchCase: false,
  matchWholeWord: false,
  formatting: {},
  mode: null,         // "text" | "format"
  dirty: true,        // הקלט השתנה מאז החיפוש האחרון
  busy: false,
  cancel: false
};

let queue = Promise.resolve();
/** מריץ פעולות מול Word אחת אחרי השנייה (Enter מהיר לא יתנגש) */
function enqueue(fn) {
  queue = queue.then(fn).catch((e) => handleError(e));
  return queue;
}

/* =========================================================
 * 4. UI helpers
 * ========================================================= */

const $ = (id) => document.getElementById(id);

function setStatus(msg, kind) {
  const el = $("status");
  el.textContent = msg || "";
  el.className = "status" + (kind ? " " + kind : "");
}

function setBusy(busy, label) {
  state.busy = busy;
  document.body.classList.toggle("busy", busy);
  ["btnSearch", "btnTake", "btnSimilar", "btnClear"].forEach((id) => { $(id).disabled = busy; });
  $("btnCancel").hidden = !busy;
  if (busy && label) setStatus(label, "info");
}

function handleError(e) {
  console.error(e);
  setBusy(false);
  let msg = (e && e.message) ? e.message : String(e);
  if (e && e.debugInfo && e.debugInfo.errorLocation) msg += " (" + e.debugInfo.errorLocation + ")";
  setStatus("שגיאה: " + msg, "error");
}

function fillSelect(sel, options) {
  sel.innerHTML = "";
  options.forEach(([v, label]) => {
    const o = document.createElement("option");
    o.value = v; o.textContent = label;
    sel.appendChild(o);
  });
}

function ensureOption(sel, value, label) {
  if (![...sel.options].some((o) => o.value === value)) {
    const o = document.createElement("option");
    o.value = value; o.textContent = label;
    sel.appendChild(o);
  }
  sel.value = value;
}

const FIELDS = {
  bold:           { chk: "chkBold",      val: "valBold" },
  italic:         { chk: "chkItalic",    val: "valItalic" },
  underline:      { chk: "chkUnderline", val: "valUnderline" },
  name:           { chk: "chkFont",      val: "valFont" },
  size:           { chk: "chkSize",      val: "valSize" },
  color:          { chk: "chkColor",     val: "valColor" },
  highlightColor: { chk: "chkHighlight", val: "valHighlight" }
};

/** קורא את קריטריוני העיצוב מהממשק. זורק שגיאה אם ערך לא תקין. */
function readCriteriaFromUI() {
  const c = {};
  if ($("chkBold").checked) c.bold = $("valBold").value === "true";
  if ($("chkItalic").checked) c.italic = $("valItalic").value === "true";
  if ($("chkUnderline").checked) c.underline = $("valUnderline").value;
  if ($("chkFont").checked) {
    const f = $("valFont").value.trim();
    if (!f) throw new Error("יש להזין שם גופן.");
    c.name = f;
  }
  if ($("chkSize").checked) {
    const n = Number($("valSize").value);
    if (!Number.isInteger(n) || n < 1 || n > 1638) throw new Error("גודל גופן חייב להיות מספר שלם.");
    c.size = n;
  }
  if ($("chkColor").checked) c.color = normalizeHex($("valColor").value);
  if ($("chkHighlight").checked) c.highlightColor = $("valHighlight").value;
  return c;
}

function markDirty() { state.dirty = true; }

function updateCounter() {
  const n = state.results.length;
  $("counter").textContent = n ? (state.currentIndex + 1) + " מתוך " + n : "0 מתוך 0";
  $("btnPrev").disabled = n === 0;
  $("btnNext").disabled = n === 0;
}

function renderResults() {
  const list = $("results");
  list.innerHTML = "";
  const frag = document.createDocumentFragment();
  state.results.forEach((r, i) => {
    const li = document.createElement("li");
    li.dataset.index = i;
    li.setAttribute("role", "option");
    const num = document.createElement("span");
    num.className = "num"; num.textContent = i + 1;
    const txt = document.createElement("span");
    txt.className = "txt"; txt.textContent = r.text;
    li.append(num, txt);
    if (r.para !== undefined) {
      const meta = document.createElement("span");
      meta.className = "meta"; meta.textContent = "פסקה " + (r.para + 1);
      li.appendChild(meta);
    }
    frag.appendChild(li);
  });
  list.appendChild(frag);
  $("resultsEmpty").hidden = state.results.length > 0 || state.mode === null;
  highlightCurrentInList();
  updateCounter();
}

/** גולל רק בתוך רשימת התוצאות – החלונית עצמה לא זזה */
function scrollListTo(li) {
  const list = $("results");
  const top = li.offsetTop - list.offsetTop;
  const bottom = top + li.offsetHeight;
  if (top < list.scrollTop) list.scrollTop = top;
  else if (bottom > list.scrollTop + list.clientHeight) list.scrollTop = bottom - list.clientHeight;
}

/** אחרי חיפוש חדש: מעלה פעם אחת את החלונית כך שהניווט והתוצאות גלויים */
function revealResultsOnce() {
  const nav = document.querySelector(".nav");
  const card = document.querySelector(".results-card");
  if (!nav || !card) return;
  const rect = card.getBoundingClientRect();
  if (rect.bottom > window.innerHeight || rect.top < 0) {
    window.scrollTo({ top: window.scrollY + nav.getBoundingClientRect().top - 8, behavior: "smooth" });
  }
}

function highlightCurrentInList() {
  const list = $("results");
  const prev = list.querySelector("li.current");
  if (prev) { prev.classList.remove("current"); prev.removeAttribute("aria-selected"); }
  const cur = list.children[state.currentIndex];
  if (cur) {
    cur.classList.add("current");
    cur.setAttribute("aria-selected", "true");
    scrollListTo(cur);
  }
}

/* =========================================================
 * 5. קריאת עיצוב
 * ========================================================= */

/** קורא רק את המאפיינים שנטענו (קריאת מאפיין שלא נטען זורקת שגיאה ב-Office.js) */
function readFont(font, props) {
  const out = {};
  (props || PROPS).forEach((p) => { out[p] = font[p]; });
  return out;
}

/* =========================================================
 * 6. חיפוש
 * ========================================================= */

function runSearch() {
  return enqueue(async () => {
    let criteria;
    try { criteria = readCriteriaFromUI(); }
    catch (e) { setStatus(e.message, "error"); return; }

    const text = $("searchText").value;
    state.searchText = text;
    state.matchCase = $("matchCase").checked;
    state.matchWholeWord = $("matchWholeWord").checked;
    state.formatting = criteria;

    if (!text.trim() && !hasAnyFormattingCriteria(criteria)) {
      setStatus("יש להזין טקסט או לבחור לפחות מאפיין עיצוב אחד.", "warn");
      return;
    }
    if (text.length > MAX_SEARCH_LEN) {
      setStatus("טקסט החיפוש ארוך מדי (עד " + MAX_SEARCH_LEN + " תווים).", "warn");
      return;
    }

    if (text.trim()) await searchText(text, criteria);
    else await searchFormatOnly(criteria);
  });
}

/** חיפוש טקסט (ואופציונלית סינון לפי עיצוב) */
async function searchText(text, criteria) {
  setBusy(true, "מחפש…");
  const hasFmt = hasAnyFormattingCriteria(criteria);
  await Word.run(async (context) => {
    const ranges = context.document.body.search(escapeWordSearch(text), {
      matchCase: state.matchCase,
      matchWholeWord: state.matchWholeWord
    });
    ranges.load(hasFmt ? buildFontLoadString(criteria, "items/") : "items/text");
    await context.sync();

    const results = [];
    ranges.items.forEach((r, i) => {
      if (results.length >= MAX_RESULTS) return;
      if (hasFmt && !formattingMatches(readFont(r.font, Object.keys(criteria)), criteria)) return;
      results.push({ text: r.text, locator: { kind: "search", index: i } });
    });
    finishSearch("text", results);
  });
  setBusy(false);
  if (state.results.length) await selectCurrentResult();
}

/** חיפוש לפי עיצוב בלבד – מעבר על המילים במסמך */
async function searchFormatOnly(criteria) {
  state.cancel = false;
  setBusy(true, "סורק את המסמך…");
  const results = [];
  const matches = [];
  let truncated = false;

  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    // פסקאות שאינן ריקות בלבד
    const nonEmpty = [];
    paragraphs.items.forEach((p, i) => { if (p.text && p.text.trim()) nonEmpty.push(i); });

    for (let start = 0; start < nonEmpty.length; start += PARAGRAPH_BATCH) {
      if (state.cancel) break;
      const batch = nonEmpty.slice(start, start + PARAGRAPH_BATCH);

      // שלב א: פירוק למילים + טעינת עיצוב – sync אחד לכל המנה
      const collections = batch.map((pi) => {
        const col = paragraphs.items[pi].getTextRanges(WORD_DELIMITERS, true);
        col.load(buildFontLoadString(criteria, "items/"));
        return col;
      });
      await context.sync();

      // שלב ב: מילים שמכילות פיסוק בקצוות ועיצוב "מעורב" –
      // בודקים את ליבת המילה בלבד (sync אחד נוסף לכל המנה)
      const pending = [];
      const candidates = [];
      batch.forEach((pi, bi) => {
        collections[bi].items.forEach((tok, ti) => {
          const core = extractCore(tok.text);
          if (!core) return;
          const font = readFont(tok.font, Object.keys(criteria));
          const entry = { para: pi, tokenIndex: ti, core, raw: tok.text };
          if (formattingMatches(font, criteria)) {
            candidates.push(entry);
          } else if (core !== tok.text.trim() && core.length <= MAX_SEARCH_LEN && hasMixedCriteria(font, criteria)) {
            const coreRange = tok.search(escapeWordSearch(core), { matchCase: true }).getFirstOrNullObject();
            coreRange.load(buildFontLoadString(criteria, ""));
            pending.push({ entry, coreRange });
            candidates.push(entry);
            entry.pending = pending[pending.length - 1];
          }
        });
      });
      if (pending.length) await context.sync();

      for (const c of candidates) {
        if (c.pending) {
          const cr = c.pending.coreRange;
          if (cr.isNullObject || !formattingMatches(readFont(cr.font, Object.keys(criteria)), criteria)) continue;
          // פיסוק בקצה בעיצוב שונה – שובר את הרצף, כמו ב-Word
          const t = c.raw.trim();
          c.breakBefore = !t.startsWith(c.core);
          c.breakAfter = !t.endsWith(c.core);
        }
        matches.push(c);
      }
      // פסקאות המנה הסתיימו – אפשר לאחד אותן לתוצאות
      for (const run of groupRuns(matches.splice(0))) {
        if (results.length >= MAX_RESULTS) { truncated = true; break; }
        const items = run.items;
        results.push({
          text: items.map((x) => x.core).join(" "),
          para: run.para,
          locator: {
            kind: "run", para: run.para,
            from: items[0].tokenIndex, to: items[items.length - 1].tokenIndex,
            raws: items.map((x) => x.raw),
            firstCore: items[0].core, lastCore: items[items.length - 1].core
          }
        });
      }
      if (truncated) break;

      const pct = Math.round(Math.min(1, (start + batch.length) / nonEmpty.length) * 100);
      setStatus("סורק את המסמך… " + pct + "% (" + results.length + " תוצאות)", "info");
    }
  });

  const cancelled = state.cancel;
  finishSearch("format", results);
  setBusy(false);
  if (cancelled) setStatus("הסריקה בוטלה. מוצגות " + results.length + " תוצאות שנמצאו עד כה.", "warn");
  if (truncated) setStatus("נמצאו יותר מ-" + MAX_RESULTS + " תוצאות; מוצגות הראשונות בלבד.", "warn");
  if (state.results.length) await selectCurrentResult(cancelled || truncated);
}

function finishSearch(mode, results) {
  state.mode = mode;
  state.results = results;
  state.currentIndex = results.length ? 0 : -1;
  state.dirty = false;
  renderResults();
  $("results").scrollTop = 0;
  if (!results.length) setStatus("לא נמצאו תוצאות.", "warn");
  else { setStatus("", ""); revealResultsOnce(); }
}

/* =========================================================
 * 7. ניווט ובחירה
 * ========================================================= */

function navigate(step) {
  return enqueue(async () => {
    const n = state.results.length;
    if (!n) return;
    state.currentIndex = ((state.currentIndex + step) % n + n) % n; // מעגלי
    highlightCurrentInList();
    updateCounter();
    await selectCurrentResult();
  });
}

function goTo(index) {
  return enqueue(async () => {
    if (index < 0 || index >= state.results.length) return;
    state.currentIndex = index;
    highlightCurrentInList();
    updateCounter();
    await selectCurrentResult();
  });
}

/**
 * בוחר את התוצאה הנוכחית ב-Word.
 * אובייקטי Range הם Proxy הקשורים ל-context, לכן ה-Range מאותר מחדש
 * בתוך Word.run לפי ה-locator, ונבחר באמצעות range.select().
 */
async function selectCurrentResult(keepStatus) {
  const r = state.results[state.currentIndex];
  if (!r) return;
  const loc = r.locator;

  await Word.run(async (context) => {
    let range = null;
    const body = context.document.body;

    if (loc.kind === "search") {
      const ranges = body.search(escapeWordSearch(state.searchText), {
        matchCase: state.matchCase, matchWholeWord: state.matchWholeWord
      });
      ranges.load("items");
      await context.sync();
      range = ranges.items[loc.index] || null;
    } else {
      const paragraphs = body.paragraphs;
      paragraphs.load("items/style");
      await context.sync();
      const p = paragraphs.items[loc.para];
      if (p) {
        const tokens = p.getTextRanges(WORD_DELIMITERS, true);
        tokens.load("items/text");
        await context.sync();
        // המסמך לא השתנה? כל המילים ברצף עדיין במקומן
        const same = loc.raws.every((raw, k) => {
          const t = tokens.items[loc.from + k];
          return t && t.text === raw;
        });
        if (same) {
          const coreOf = (tok, raw, core) => (core !== raw)
            ? tok.search(escapeWordSearch(core), { matchCase: true }).getFirstOrNullObject()
            : tok;
          const first = coreOf(tokens.items[loc.from], loc.raws[0], loc.firstCore);
          const last = loc.to === loc.from ? first
            : coreOf(tokens.items[loc.to], loc.raws[loc.raws.length - 1], loc.lastCore);
          first.load("isNullObject");
          if (last !== first) last.load("isNullObject");
          await context.sync();
          if (!first.isNullObject && !last.isNullObject) {
            range = (last === first) ? first : first.expandTo(last);
          }
        }
      }
    }

    if (!range) { reportStale(); return; }
    range.load("text,isNullObject");
    await context.sync();
    if (range.isNullObject) { reportStale(); return; }
    if (loc.kind === "search" && !textEquivalent(range.text, r.text)) { reportStale(); return; }

    range.select("Select");
    await context.sync();
    if (!keepStatus) setStatus("", "");
  });
}

function textEquivalent(a, b) {
  return String(a).trim() === String(b).trim();
}

function reportStale() {
  setStatus("המסמך השתנה מאז החיפוש – לחצו \"חפש\" שוב לרענון התוצאות.", "warn");
  state.dirty = true;
}

/* =========================================================
 * 8. עיצוב מהבחירה
 * ========================================================= */

/** קורא את עיצוב הבחירה הנוכחית. מחזיר אובייקט font או null. */
async function readSelectionFont() {
  let result = null;
  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    sel.font.load(PROPS.join(","));
    await context.sync();

    let font = readFont(sel.font);
    let text = sel.text || "";
    const trimmed = text.trim();

    // לחיצה כפולה בוחרת גם את הרווח שאחרי המילה – קוראים את המילה בלבד
    if (trimmed && trimmed !== text && trimmed.length <= MAX_SEARCH_LEN) {
      const inner = sel.search(escapeWordSearch(trimmed), { matchCase: true }).getFirstOrNullObject();
      inner.load("isNullObject");
      inner.font.load(PROPS.join(","));
      await context.sync();
      if (!inner.isNullObject) font = readFont(inner.font);
    }
    result = { font, text: trimmed };
  });
  return result;
}

/** ממלא את שדות הממשק; מחזיר רשימת מאפיינים מעורבים */
function applyFontToUI(font, opts) {
  const mixed = [];
  const anyChecked = PROPS.some((p) => $(FIELDS[p].chk).checked);

  PROPS.forEach((p) => {
    const v = font[p];
    const chk = $(FIELDS[p].chk);
    const val = $(FIELDS[p].val);
    if (isMixed(p, v)) {
      mixed.push(p);
      chk.checked = false;
      return;
    }
    switch (p) {
      case "bold":
      case "italic":
        val.value = v ? "true" : "false"; break;
      case "underline":
        ensureOption(val, String(v), String(v)); break;
      case "name":
        val.value = v; break;
      case "size":
        val.value = Math.round(Number(v)); break;
      case "color": {
        const hex = normalizeHex(v);
        if (hex) val.value = hex.toLowerCase(); else { mixed.push(p); chk.checked = false; return; }
        break;
      }
      case "highlightColor": {
        const h = normalizeHighlight(v);
        if (!h) { mixed.push(p); chk.checked = false; return; }
        ensureOption(val, h, h);
        break;
      }
    }
    if (opts && opts.checkAllIfNone && !anyChecked) chk.checked = true;
  });
  markDirty();
  return mixed;
}

const PROP_LABELS = {
  bold: "מודגש", italic: "נטוי", underline: "קו תחתון", name: "גופן",
  size: "גודל", color: "צבע טקסט", highlightColor: "צבע הדגשה"
};

function takeFormattingFromSelection() {
  return enqueue(async () => {
    setBusy(true, "קורא עיצוב מהבחירה…");
    const sel = await readSelectionFont();
    setBusy(false);
    const mixed = applyFontToUI(sel.font, { checkAllIfNone: false });
    if (mixed.length) {
      setStatus("העיצוב נטען. מאפיינים מעורבים בבחירה (לא סומנו): " + mixed.map((p) => PROP_LABELS[p]).join(", "), "info");
    } else {
      setStatus("העיצוב נטען מהבחירה. סמנו את המאפיינים שברצונכם לחפש.", "info");
    }
  });
}

function findSimilarFormatting() {
  return enqueue(async () => {
    setBusy(true, "קורא עיצוב מהבחירה…");
    const sel = await readSelectionFont();
    setBusy(false);
    if (!sel.text) {
      setStatus("יש לבחור מילה במסמך תחילה.", "warn");
      return;
    }
    // אם המשתמש עוד לא בחר מאפיינים – משווים את כל המאפיינים שאינם מעורבים
    applyFontToUI(sel.font, { checkAllIfNone: true });

    let criteria;
    try { criteria = readCriteriaFromUI(); }
    catch (e) { setStatus(e.message, "error"); return; }
    if (!hasAnyFormattingCriteria(criteria)) {
      setStatus("לבחירה אין מאפיין עיצוב אחיד להשוואה.", "warn");
      return;
    }
    state.searchText = "";
    state.formatting = criteria;
    await searchFormatOnly(criteria);
  });
}

function clearFormatting() {
  PROPS.forEach((p) => { $(FIELDS[p].chk).checked = false; });
  $("valBold").value = "true";
  $("valItalic").value = "true";
  $("valUnderline").value = "__any__";
  $("valFont").value = "";
  $("valSize").value = "12";
  $("valColor").value = "#000000";
  $("valHighlight").value = "#FFFF00";
  markDirty();
  setStatus("", "");
}

/* =========================================================
 * 9. אתחול
 * ========================================================= */

function wireUI() {
  fillSelect($("valUnderline"), UNDERLINE_OPTIONS);
  fillSelect($("valHighlight"), HIGHLIGHT_OPTIONS);
  clearFormatting();

  $("btnSearch").addEventListener("click", runSearch);
  $("btnNext").addEventListener("click", () => navigate(+1));
  $("btnPrev").addEventListener("click", () => navigate(-1));
  $("btnTake").addEventListener("click", takeFormattingFromSelection);
  $("btnSimilar").addEventListener("click", findSimilarFormatting);
  $("btnClear").addEventListener("click", clearFormatting);
  $("btnCancel").addEventListener("click", () => { state.cancel = true; });

  $("results").addEventListener("click", (e) => {
    const li = e.target.closest("li");
    if (li) goTo(Number(li.dataset.index));
  });

  // כל שינוי בקלט מסמן שצריך חיפוש חדש
  document.querySelectorAll("#searchText, #matchCase, #matchWholeWord, .fmt input, .fmt select")
    .forEach((el) => {
      el.addEventListener("input", markDirty);
      el.addEventListener("change", markDirty);
    });

  // סימון אוטומטי של checkbox כשמשנים ערך
  PROPS.forEach((p) => {
    const val = $(FIELDS[p].val);
    const handler = () => { $(FIELDS[p].chk).checked = true; };
    val.addEventListener("change", handler);
    if (val.tagName === "INPUT" && val.type !== "color") val.addEventListener("input", handler);
  });

  // Enter = הבא, Shift+Enter = הקודם (או חיפוש אם הקלט השתנה)
  $("searchText").addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    if (state.busy) return;
    if (state.dirty || !state.results.length) runSearch();
    else navigate(e.shiftKey ? -1 : +1);
  });

  renderResults();
}

if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    wireUI();
    if (info.host !== Office.HostType.Word) {
      setStatus("התוסף מיועד ל-Microsoft Word.", "error");
      return;
    }
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      setStatus("גרסת Word אינה תומכת ב-WordApi 1.3. יש לעדכן את Office.", "error");
    }
    $("searchText").focus();
  });
}
