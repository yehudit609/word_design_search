// בדיקות ללוגיקה הטהורה של taskpane.js:  node tests/logic.test.js
"use strict";
const assert = require("assert");
const L = require("../src/taskpane.js");

let passed = 0;
function t(name, fn) {
  try { fn(); passed++; }
  catch (e) { console.error("FAIL:", name, "\n ", e.message); process.exitCode = 1; }
}

const base = { bold: true, italic: false, underline: "None", name: "Arial", size: 14, color: "#000000", highlightColor: null };

// extractCore – סימני פיסוק בעברית ובאנגלית
t("core comma", () => assert.strictEqual(L.extractCore("שלום,"), "שלום"));
t("core period", () => assert.strictEqual(L.extractCore("שלום."), "שלום"));
t("core colon", () => assert.strictEqual(L.extractCore("שלום:"), "שלום"));
t("core bang", () => assert.strictEqual(L.extractCore("שלום!"), "שלום"));
t("core quotes", () => assert.strictEqual(L.extractCore("״הסכם״"), "הסכם"));
t("core gershayim inside", () => assert.strictEqual(L.extractCore("עו״ד"), "עו״ד"));
t("core geresh inside", () => assert.strictEqual(L.extractCore("צ׳יפס"), "צ׳יפס"));
t("core maqaf", () => assert.strictEqual(L.extractCore("בית־"), "בית"));
t("core dashes", () => assert.strictEqual(L.extractCore("—מילה–"), "מילה"));
t("core english", () => assert.strictEqual(L.extractCore("(Hello)."), "Hello"));
t("core pure punct", () => assert.strictEqual(L.extractCore("…"), ""));
t("core spaces", () => assert.strictEqual(L.extractCore("  מחיר  "), "מחיר"));

// formattingMatches – כל מאפיין לבד
t("bold only", () => assert.ok(L.formattingMatches(base, { bold: true })));
t("bold false mismatch", () => assert.ok(!L.formattingMatches(base, { bold: false })));
t("italic", () => assert.ok(L.formattingMatches(base, { italic: false })));
t("underline any rejects None", () => assert.ok(!L.formattingMatches(base, { underline: "__any__" })));
t("underline any accepts Double", () => assert.ok(L.formattingMatches({ ...base, underline: "Double" }, { underline: "__any__" })));
t("underline exact", () => assert.ok(L.formattingMatches({ ...base, underline: "Single" }, { underline: "single" })));
t("font case-insensitive", () => assert.ok(L.formattingMatches(base, { name: " arial " })));
t("font mismatch", () => assert.ok(!L.formattingMatches(base, { name: "David" })));
t("size int", () => assert.ok(L.formattingMatches(base, { size: 14 })));
t("size 14.5 != 14", () => assert.ok(!L.formattingMatches({ ...base, size: 14.5 }, { size: 14 })));
t("color case", () => assert.ok(L.formattingMatches({ ...base, color: "#ff0000" }, { color: "#FF0000" })));
t("highlight none", () => assert.ok(L.formattingMatches(base, { highlightColor: "none" })));
t("highlight name→hex", () => assert.ok(L.formattingMatches({ ...base, highlightColor: "Yellow" }, { highlightColor: "#FFFF00" })));
t("highlight hex", () => assert.ok(L.formattingMatches({ ...base, highlightColor: "#ffff00" }, { highlightColor: "#FFFF00" })));
t("highlight any", () => assert.ok(L.formattingMatches({ ...base, highlightColor: "Turquoise" }, { highlightColor: "__any__" })));
t("highlight any rejects none", () => assert.ok(!L.formattingMatches(base, { highlightColor: "__any__" })));

// שילובים
t("bold+font", () => assert.ok(L.formattingMatches(base, { bold: true, name: "Arial" })));
t("bold+size", () => assert.ok(L.formattingMatches(base, { bold: true, size: 14 })));
t("font+size mismatch", () => assert.ok(!L.formattingMatches(base, { name: "Arial", size: 12 })));
t("color+highlight", () => assert.ok(L.formattingMatches({ ...base, highlightColor: "Yellow" }, { color: "#000000", highlightColor: "#FFFF00" })));
t("all props", () => assert.ok(L.formattingMatches(base, { bold: true, italic: false, underline: "None", name: "Arial", size: 14, color: "#000000", highlightColor: "none" })));
t("unselected props ignored", () => assert.ok(L.formattingMatches({ ...base, color: "#123456" }, { bold: true, size: 14 })));

// עיצוב מעורב
t("mixed bold rejected", () => assert.ok(!L.formattingMatches({ ...base, bold: null }, { bold: true })));
t("mixed underline rejected", () => assert.ok(!L.formattingMatches({ ...base, underline: "Mixed" }, { underline: "__any__" })));
t("mixed size rejected", () => assert.ok(!L.formattingMatches({ ...base, size: null }, { size: 14 })));
t("mixed ignored if not criterion", () => assert.ok(L.formattingMatches({ ...base, size: null }, { bold: true })));
t("hasMixedCriteria", () => assert.ok(L.hasMixedCriteria({ ...base, bold: null }, { bold: true })));

// עזר
t("escape caret", () => assert.strictEqual(L.escapeWordSearch("a^b"), "a^^b"));
t("load string", () => assert.strictEqual(L.buildFontLoadString({ bold: true, size: 12 }, "items/"), "items/text,items/font/bold,items/font/size"));
t("no criteria", () => assert.ok(!L.hasAnyFormattingCriteria({})));
t("delimiters exclude quotes", () => assert.ok(!L.WORD_DELIMITERS.includes("״") && !L.WORD_DELIMITERS.includes("'")));

console.log(passed + " tests passed" + (process.exitCode ? " (with failures)" : ""));
