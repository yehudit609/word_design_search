// בדיקת UI מקצה לקצה מול Word מדומה:  node tests/ui.test.js  (דורש playwright)
"use strict";
const path = require("path");
const fs = require("fs");
const assert = require("assert");
const { chromium } = require("playwright");

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 340, height: 900 } });
  const mock = fs.readFileSync(path.join(__dirname, "mock-office.js"), "utf8");
  await page.route("**/office.js", (r) => r.fulfill({ contentType: "text/javascript", body: mock }));
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto("file://" + path.join(__dirname, "..", "src", "taskpane.html").replace(/\\/g, "/"));
  await page.waitForTimeout(100);

  const counter = () => page.textContent("#counter");
  const status = () => page.textContent("#status");
  const lastSel = () => page.evaluate(() => window.__selected[window.__selected.length - 1]);
  const idle = () => page.waitForFunction(() => !document.body.classList.contains("busy")).then(() => page.waitForTimeout(50));
  let n = 0;
  const ok = (name, cond) => { assert.ok(cond, name); n++; };

  // תרחיש 7 – אין קריטריונים
  await page.click("#btnSearch"); await idle();
  ok("no criteria msg", (await status()).includes("יש להזין טקסט או לבחור לפחות מאפיין עיצוב אחד."));

  // תרחיש 1 – חיפוש טקסט, בחירה, הבא
  await page.fill("#searchText", "חוזה");
  await page.click("#btnSearch"); await idle();
  ok("text count", (await counter()) === "1 מתוך 3");
  ok("first selected", (await lastSel()).meta.pi === 0);
  await page.click("#btnNext"); await idle();
  ok("next", (await counter()) === "2 מתוך 3" && (await lastSel()).meta.pi === 2);

  // תרחיש 5 – מעגלי + Enter / Shift+Enter
  await page.focus("#searchText");
  await page.keyboard.press("Enter"); await idle();
  ok("enter next", (await counter()) === "3 מתוך 3");
  await page.keyboard.press("Enter"); await idle();
  ok("wrap to 1", (await counter()) === "1 מתוך 3");
  await page.keyboard.press("Shift+Enter"); await idle();
  ok("wrap to last", (await counter()) === "3 מתוך 3");

  // לחיצה ישירה על תוצאה
  await page.click("#results li:nth-child(2)"); await idle();
  ok("click result", (await counter()) === "2 מתוך 3" && (await page.$eval("#results li.current .num", (e) => e.textContent)) === "2");

  // מילה שלמה / אותיות
  await page.check("#matchWholeWord");
  await page.click("#btnSearch"); await idle();
  ok("whole word excludes החוזה", (await counter()) === "1 מתוך 2");
  await page.uncheck("#matchWholeWord");
  await page.fill("#searchText", "Price"); await page.check("#matchCase");
  await page.click("#btnSearch"); await idle();
  ok("match case", (await counter()) === "1 מתוך 1");
  await page.uncheck("#matchCase");

  // תרחיש 6 – אין תוצאות
  await page.fill("#searchText", "לאקיים");
  await page.click("#btnSearch"); await idle();
  ok("no results", (await status()).includes("לא נמצאו תוצאות.") && (await counter()) === "0 מתוך 0");

  // תרחיש 2 – Bold בלבד (כולל "המוכר," עם פסיק לא מודגש)
  await page.fill("#searchText", "");
  await page.check("#chkBold");
  await page.click("#btnSearch"); await idle();
  const boldTexts = await page.$$eval("#results .txt", (els) => els.map((e) => e.textContent));
  ok("bold only results", JSON.stringify(boldTexts) === JSON.stringify(["הסכם", "הצדדים", "המוכר", "מחיר", "הקונה"]));
  await page.click("#results li:nth-child(3)"); await idle();
  ok("punct word selects core", (await lastSel()).text === "המוכר");

  // טקסט + עיצוב
  await page.fill("#searchText", "מחיר");
  await page.check("#chkSize"); await page.fill("#valSize", "14");
  await page.click("#btnSearch"); await idle();
  ok("text+format", (await counter()) === "1 מתוך 1");
  await page.fill("#valSize", "12");
  await page.click("#btnSearch"); await idle();
  ok("text+format filtered out", (await counter()) === "0 מתוך 0");

  // תרחיש 3 – קח עיצוב מהבחירה (המילה "הסכם" + רווח)
  await page.click("#btnClear"); await page.fill("#searchText", "");
  await page.evaluate(() => { window.__selectionToken = [0, 0]; });
  await page.click("#btnTake"); await idle();
  ok("take font", (await page.inputValue("#valFont")) === "Arial");
  ok("take size", (await page.inputValue("#valSize")) === "14");
  ok("take bold value", (await page.inputValue("#valBold")) === "true");
  ok("take highlight", (await page.inputValue("#valHighlight")) === "#FFFF00");
  ok("take does not check boxes", !(await page.isChecked("#chkBold")));

  // תרחיש 4 – מצא מילים בעיצוב דומה (כל המאפיינים) → הסכם, הצדדים, המוכר, הקונה (לא "מחיר" – בלי הדגשה)
  await page.click("#btnSimilar"); await idle();
  const sim = await page.$$eval("#results .txt", (els) => els.map((e) => e.textContent));
  ok("similar results", JSON.stringify(sim) === JSON.stringify(["הסכם", "הצדדים", "המוכר", "הקונה"]));
  ok("similar first selected", (await lastSel()).text === "הסכם");

  // "דומה" לפי המאפיינים שסומנו בלבד: בלי הדגשה → גם "מחיר"
  await page.uncheck("#chkHighlight");
  await page.click("#btnSimilar"); await idle();
  const sim2 = await page.$$eval("#results .txt", (els) => els.map((e) => e.textContent));
  ok("similar respects checkboxes", sim2.includes("מחיר") && sim2.length === 5);

  await page.screenshot({ path: path.join(__dirname, "taskpane-screenshot.png"), fullPage: true });
  ok("no page errors", errors.length === 0 || (console.error(errors), false));
  await browser.close();
  console.log(n + " UI checks passed");
})().catch((e) => { console.error("FAIL:", e.message); process.exit(1); });
