// Mock מינימלי של Office.js/Word לבדיקות UI אוטומטיות (Playwright).
// מסמך מדומה: פסקאות → מילים, לכל מילה עיצוב.
(function () {
  const F = (o) => Object.assign({ bold: false, italic: false, underline: "None", name: "David", size: 12, color: "#000000", highlightColor: null }, o);
  const B = { bold: true, name: "Arial", size: 14, highlightColor: "Yellow" };

  // [טקסט גולמי, עיצוב גולמי, עיצוב הליבה (אם שונה)]
  const doc = [
    [["הסכם", F(B)], ["בין", F()], ["הצדדים", F(B)], ["החוזה", F()]],
    [],
    [["המוכר,", F({ ...B, bold: null, highlightColor: "" }), F(B)], ["מחיר", F({ bold: true, name: "Arial", size: 14 })], ["חוזה", F()]],
    [["הקונה", F(B)], ["חוזה.", F()], ["Price", F({ italic: true })], ["price", F()]]
  ];
  window.__selected = [];
  window.__selectionToken = [0, 0];

  let uid = 0;
  function mkRange(text, font, meta) {
    if (font && !font.load) Object.defineProperty(font, "load", { value() { return font; }, enumerable: false });
    const r = { id: ++uid, text, font, isNullObject: false, meta,
      load() { return r; },
      select() { window.__selected.push({ text: r.text, meta: r.meta }); },
      search(q) {
        const core = q.replace(/\^\^/g, "^");
        const found = r.text.includes(core) ? [mkRange(core, r.coreFont || r.font, r.meta)] : [];
        return mkCol(found);
      }
    };
    return r;
  }
  function mkCol(items) {
    return { items, load() { return this; },
      getFirstOrNullObject() { return items[0] || { isNullObject: true, load() { return this; }, font: { load() {} } }; } };
  }
  function tokenRange(pi, ti) {
    const [text, font, coreFont] = doc[pi][ti];
    const r = mkRange(text, font, { pi, ti });
    r.coreFont = coreFont;
    return r;
  }

  const body = {
    get paragraphs() {
      return mkCol(doc.map((p, pi) => ({
        text: p.map((t) => t[0]).join(" "),
        style: "Normal",
        getTextRanges() { return mkCol(p.map((_, ti) => tokenRange(pi, ti))); }
      })));
    },
    search(q, opts) {
      const out = [];
      doc.forEach((p, pi) => p.forEach(([text], ti) => {
        const words = text.replace(/[.,]/g, "");
        const a = opts.matchCase ? words : words.toLowerCase();
        const b = opts.matchCase ? q : q.toLowerCase();
        const hit = opts.matchWholeWord ? a === b : a.includes(b);
        if (hit) {
          const t = tokenRange(pi, ti);
          const r = mkRange(words.substr(a.indexOf(b), q.length), t.coreFont || t.font, { pi, ti });
          out.push(r);
        }
      }));
      return mkCol(out);
    }
  };

  const context = {
    document: {
      body,
      getSelection() {
        const [pi, ti] = window.__selectionToken;
        const t = tokenRange(pi, ti);
        // לחיצה כפולה: המילה + רווח (רווח לא מודגש → מעורב)
        const sel = mkRange(t.text + " ", F({ ...t.font, bold: null }), { pi, ti });
        sel.search = () => mkCol([mkRange(t.text, t.coreFont || t.font, { pi, ti })]);
        return sel;
      }
    },
    sync() { return Promise.resolve(); }
  };

  window.Word = { run: (cb) => Promise.resolve().then(() => cb(context)) };
  window.Office = {
    HostType: { Word: "Word" },
    context: { requirements: { isSetSupported: () => true } },
    onReady(cb) { setTimeout(() => cb({ host: "Word" }), 0); return Promise.resolve(); }
  };
})();
