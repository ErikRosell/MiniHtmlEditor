// End-to-end tests for Mini HTML Editor, run in headless Chromium.
//   npm install && npm test
import { fileURLToPath, pathToFileURL } from 'node:url';
import { execSync } from 'node:child_process';
import path from 'node:path';

let chromium;
try {
  ({ chromium } = await import('playwright'));
} catch {
  const globalRoot = execSync('npm root -g').toString().trim();
  ({ chromium } = await import(pathToFileURL(path.join(globalRoot, 'playwright', 'index.mjs')).href));
}

const here = path.dirname(fileURLToPath(import.meta.url));
const APP = pathToFileURL(path.join(here, '..', 'index.html')).href;

const P = 'margin:0 0 12px 0;';
const P0 = 'margin:0;';
const H = 'margin:0 0 6px 0;font-size:16px;font-weight:600;color:var(--muBlue);';
const LI = 'margin:0 0 6px 0;';
const UL = 'margin:0 0 12px 0;padding-left:20px;';

// The example output from the original specification.
const EXAMPLE = '<div style="max-width:760px;font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#222;"><p style="margin:0 0 12px 0;">In this test, you will see a series of numbers. The numbers follow a rule based on addition, subtraction, multiplication, division, or a combination of these operations.</p><div style="margin:0 0 14px 0;padding:12px 14px;background:#f7f9fc;border-left:4px solid var(--muBlue);border-radius:4px;"><div style="margin:0 0 6px 0;font-size:16px;font-weight:600;color:var(--muBlue);">Your task</div><p style="margin:0;">Find the next number in the series, enter it in the box to the right, and click Next.</p></div><div style="margin:0 0 14px 0;"><div style="margin:0 0 6px 0;font-size:16px;font-weight:600;color:var(--muBlue);">Important rules</div><ul style="margin:0;padding-left:20px;"><li style="margin:0 0 6px 0;">Look for the simplest rule in the series, not the most complex one.</li><li style="margin:0 0 6px 0;">The length of the number series may vary from one task to another.</li><li style="margin:0 0 6px 0;">You only need to enter the very next logical number in the series.</li><li style="margin:0 0 6px 0;">Use the numerical keys to enter your answer.</li><li style="margin:0 0 6px 0;">Each task has only one correct answer.</li><li style="margin:0;">You may change your answer by entering a different number before clicking Next.</li></ul></div><div style="margin:0 0 14px 0;"><div style="margin:0 0 6px 0;font-size:16px;font-weight:600;color:var(--muBlue);">Time and navigation</div><ul style="margin:0;padding-left:20px;"><li style="margin:0 0 6px 0;">There are 12 tasks, and you have 10 minutes to complete as many as possible.</li><li style="margin:0 0 6px 0;">You may complete the tasks in any order.</li><li style="margin:0 0 6px 0;">To move between tasks, click the task number at the bottom of the page.</li><li style="margin:0 0 6px 0;">A blank circle means no answer has been given. A filled circle means an answer has been entered.</li><li style="margin:0 0 6px 0;">You may skip a task and return to it later if time remains.</li><li style="margin:0;">A clock will show the remaining time and will turn red when 2 minutes remain.</li></ul></div><div><div style="margin:0 0 6px 0;font-size:16px;font-weight:600;color:var(--muBlue);">Before you begin</div><p style="margin:0 0 8px 0;">On the next pages, you will see three examples before the timed test begins.</p><p style="margin:0;">Please make sure you have read and understood the instructions before proceeding.</p></div></div>';

const WORD_HTML = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta name=Generator content="Microsoft Word 15"><style><!-- p.MsoNormal {margin:0cm;} --></style></head><body lang=SV style='tab-interval:65.2pt'>
<!--StartFragment-->
<p class=MsoNormal><b><span style='font-size:14.0pt;mso-bidi-font-size:11.0pt'>Instructions<o:p></o:p></span></b></p>
<p class=MsoNormal>In this test, you will be shown one
pair of words.<o:p></o:p></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoListParagraphCxSpFirst style='text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>There may be several possible
relationships.<o:p></o:p></p>
<p class=MsoListParagraphCxSpMiddle style='text-indent:-18.0pt;mso-list:l0 level2 lfo1'><![if !supportLists]><span style='font-family:"Courier New"'><span style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp; </span></span></span><![endif]>Nested item<o:p></o:p></p>
<p class=MsoListParagraphCxSpLast style='text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>Each task has only one correct answer.<o:p></o:p></p>
<p class=MsoListParagraphCxSpFirst style='text-indent:-18.0pt;mso-list:l1 level1 lfo2'><![if !supportLists]><span><span style='mso-list:Ignore'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>First step<o:p></o:p></p>
<p class=MsoListParagraphCxSpLast style='text-indent:-18.0pt;mso-list:l1 level1 lfo2'><![if !supportLists]><span><span style='mso-list:Ignore'>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>Second step<o:p></o:p></p>
<!--EndFragment-->
</body></html>`;

const GDOCS_HTML = '<meta charset="utf-8"><b style="font-weight:normal;" id="docs-internal-guid-1234"><h2 dir="ltr" style="line-height:1.38;margin-top:18pt;margin-bottom:6pt;"><span style="font-size:16pt;font-family:Arial;color:#000000;font-weight:400;">Rubrik från Docs</span></h2><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt;"><span style="font-size:11pt;font-family:Arial;color:#000000;font-weight:700;">Fet</span><span style="font-size:11pt;font-family:Arial;color:#000000;font-weight:400;"> och </span><span style="font-size:11pt;font-family:Arial;color:#000000;font-weight:400;font-style:italic;">kursiv</span></p><ul style="margin-top:0;margin-bottom:0;padding-inline-start:48px;"><li dir="ltr" style="list-style-type:disc;font-size:11pt;" aria-level="1"><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt;" role="presentation"><span style="font-size:11pt;">Punkt ett</span></p></li><li dir="ltr" style="list-style-type:disc;font-size:11pt;" aria-level="1"><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt;" role="presentation"><span style="font-size:11pt;">Punkt två</span></p></li></ul></b>';

/* ---------- tiny test runner ---------- */
const results = [];
let page, browser;
async function test(name, fn) {
  await fresh();
  try { await fn(); results.push({ name, ok: true }); process.stdout.write(`  ✓ ${name}\n`); }
  catch (e) { results.push({ name, ok: false, err: e }); process.stdout.write(`  ✗ ${name}\n      ${String(e.message).split('\n').join('\n      ')}\n`); }
}
function eq(actual, expected, label = '') {
  if (actual !== expected) throw new Error(`${label}\n   expected: ${expected}\n   actual:   ${actual}`);
}
function ok(cond, label) { if (!cond) throw new Error(label); }

async function fresh() {
  await page.evaluate(() => localStorage.clear());
  await page.goto(APP);
  await page.waitForFunction(() => window.MHE);
}
async function setDoc(html) {
  await page.evaluate(h => {
    const ed = document.getElementById('editor');
    ed.innerHTML = h; MHE.normalizeEditor(); MHE.hist.init(); ed.focus();
  }, html);
}
const out = (opts = { wrapper: false }) => page.evaluate(o => MHE.serializeDoc(o), opts);
async function caret(text, offset = 0, nth = 0) {
  await page.evaluate(({ text, offset, nth }) => {
    const ed = document.getElementById('editor');
    const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT); let n, k = 0;
    while ((n = w.nextNode())) { const i = n.data.indexOf(text); if (i >= 0 && k++ === nth) { ed.focus(); const r = document.createRange(); r.setStart(n, i + offset); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); return; } }
    throw new Error('text not found: ' + text);
  }, { text, offset, nth });
}
async function caretIn(selector, index = 0, where = 'start') {
  await page.evaluate(({ selector, index, where }) => {
    const ed = document.getElementById('editor');
    const el = ed.querySelectorAll(selector)[index]; ed.focus();
    const r = document.createRange(); r.selectNodeContents(el); r.collapse(where === 'start');
    getSelection().removeAllRanges(); getSelection().addRange(r);
  }, { selector, index, where });
}
async function select(from, to = from, fromOff = 0, toOff = null) {
  await page.evaluate(({ from, to, fromOff, toOff }) => {
    const ed = document.getElementById('editor');
    const find = t => { const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT); let n; while ((n = w.nextNode())) { const i = n.data.indexOf(t); if (i >= 0) return [n, i]; } throw new Error('text not found: ' + t); };
    const [a, ai] = find(from), [b, bi] = find(to);
    ed.focus();
    const r = document.createRange(); r.setStart(a, ai + fromOff); r.setEnd(b, bi + (toOff == null ? to.length : toOff));
    getSelection().removeAllRanges(); getSelection().addRange(r);
  }, { from, to, fromOff, toOff });
}
async function paste(data) {
  await page.evaluate(({ html, text }) => {
    const dt = new DataTransfer();
    if (html != null) dt.setData('text/html', html);
    if (text != null) dt.setData('text/plain', text);
    document.getElementById('editor').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, data);
}
const key = k => page.keyboard.press(k);
const type = t => page.keyboard.type(t);
const settle = () => page.waitForTimeout(60);

/* ---------- tests ---------- */
browser = await chromium.launch();
page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
const pageErrors = [];
page.on('pageerror', e => pageErrors.push(e.message));
await page.goto(APP);

console.log('\nOutput format');
await test('output is a single line with inline CSS and the wrapper', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Två</p>`);
  const html = await out({});
  ok(!html.includes('\n'), 'contains a newline');
  eq(html, `<div style="max-width:760px;font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#222;"><p style="${P}">Ett</p><p style="${P0}">Två</p></div>`);
});
await test('the example from the specification round-trips exactly (paste as HTML code)', async () => {
  await setDoc(`<p style="${P}"><br></p>`);
  await caretIn('p');
  await paste({ text: EXAMPLE });
  eq(await out({}), EXAMPLE);
});
await test('the example round-trips through "Redigera HTML"', async () => {
  await setDoc(`<p style="${P}"><br></p>`);
  await caretIn('p'); await paste({ text: EXAMPLE });
  await page.click('#code-edit');
  const pretty = await page.$eval('#code-textarea', t => t.value);
  ok(pretty.split('\n').length > 10, 'edit view should be indented over many lines');
  await page.click('#code-apply');
  eq(await out({}), EXAMPLE);
  eq(await page.$eval('#code-out', e => e.textContent), EXAMPLE, 'code panel');
});

console.log('\nEnter, blank lines and spacing');
await test('Enter in the middle of a paragraph splits it at the cursor', async () => {
  await setDoc(`<p style="${P}">Hello World</p><p style="${P}">Slut</p>`);
  await caret('Hello World', 5);
  await key('Enter');
  eq(await out(), `<p style="${P}">Hello</p><p style="${P}">World</p><p style="${P0}">Slut</p>`);
  await type('X');
  eq(await out(), `<p style="${P}">Hello</p><p style="${P}">XWorld</p><p style="${P0}">Slut</p>`, 'caret should be at start of the new paragraph');
});
await test('Enter keeps the cursor in place, not at the end of the document', async () => {
  await setDoc(`<p style="${P}">Första</p><p style="${P}">Andra</p><p style="${P}">Tredje</p>`);
  await caret('Första', 6);
  await key('Enter'); await type('Ny');
  eq(await out(), `<p style="${P}">Första</p><p style="${P}">Ny</p><p style="${P}">Andra</p><p style="${P0}">Tredje</p>`);
});
await test('Enter twice gives a visible blank line in output and preview', async () => {
  await setDoc(`<p style="${P}">Rad ett</p>`);
  await caret('Rad ett', 7);
  await key('Enter'); await key('Enter'); await type('Rad två');
  eq(await out(), `<p style="${P}">Rad ett</p><p style="${P}">&nbsp;</p><p style="${P0}">Rad två</p>`);
  await settle();
  const heights = await page.evaluate(() => Array.from(document.getElementById('preview').contentDocument.querySelectorAll('p')).map(p => p.getBoundingClientRect().height));
  ok(heights.length === 3 && heights[1] > 10, 'blank line has no height in preview: ' + heights.join(','));
});
await test('Enter at the end of a heading continues with body text', async () => {
  await setDoc(`<div style="${H}" data-preset="h-sub">Rubrik</div><p style="${P}">Slut</p>`);
  await caret('Rubrik', 6);
  await key('Enter'); await type('Brödtext');
  eq(await out(), `<div style="${H}">Rubrik</div><p style="${P}">Brödtext</p><p style="${P0}">Slut</p>`);
});
await test('Shift+Enter inserts a line break inside the paragraph', async () => {
  await setDoc(`<p style="${P}">AB</p><p style="${P}">Slut</p>`);
  await caret('AB', 1);
  await key('Shift+Enter');
  eq(await out(), `<p style="${P}">A<br>B</p><p style="${P0}">Slut</p>`);
});
await test('"↓ avstånd" sets margin-bottom on the current paragraph only', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Två</p><p style="${P}">Tre</p>`);
  await caret('Ett', 1);
  await page.selectOption('#sel-mb', '24px');
  eq(await out(), `<p style="margin:0 0 24px 0;">Ett</p><p style="${P}">Två</p><p style="${P0}">Tre</p>`);
});
await test('"↑ avstånd" and line height apply to all selected paragraphs', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Två</p><p style="${P}">Tre</p>`);
  await select('Ett', 'Två');
  await page.selectOption('#sel-mt', '20px');
  await page.selectOption('#sel-lh', '1.75');
  eq(await out(), `<p style="margin:20px 0 12px 0;line-height:1.75;">Ett</p><p style="margin:20px 0 12px 0;line-height:1.75;">Två</p><p style="${P0}">Tre</p>`);
});

console.log('\nText formatting and styles');
await test('font size on a selected word changes only that word', async () => {
  await setDoc(`<p style="${P}">Hej fina världen</p><p style="${P}">Slut</p>`);
  await select('fina');
  await page.selectOption('#sel-size', '20px');
  eq(await out(), `<p style="${P}">Hej <span style="font-size:20px;">fina</span> världen</p><p style="${P0}">Slut</p>`);
});
await test('a text preset on a selected word does not touch the paragraph', async () => {
  await setDoc(`<p style="${P}">Hej fina världen</p><p style="${P}">Slut</p>`);
  await select('fina');
  await page.click('.preset[data-id="t-blue"] .preset-apply');
  eq(await out(), `<p style="${P}">Hej <span style="color:var(--muBlue);">fina</span> världen</p><p style="${P0}">Slut</p>`);
});
await test('text color from the palette, including CSS variables', async () => {
  await setDoc(`<p style="${P}">Hej fina världen</p><p style="${P}">Slut</p>`);
  await select('världen');
  await page.click('[data-pop="color"]');
  await page.click('.popover .sw[data-v="var(--muBlue)"]');
  eq(await out(), `<p style="${P}">Hej fina <span style="color:var(--muBlue);">världen</span></p><p style="${P0}">Slut</p>`);
});
await test('removing color from part of a colored span splits it cleanly', async () => {
  await setDoc(`<p style="${P}"><span style="color:#c92a2a;">abcdef</span></p><p style="${P}">Slut</p>`);
  await select('abcdef', 'abcdef', 2, 4);
  await page.click('[data-pop="color"]');
  await page.click('.popover [data-act="none"]');
  eq(await out(), `<p style="${P}"><span style="color:#c92a2a;">ab</span>cd<span style="color:#c92a2a;">ef</span></p><p style="${P0}">Slut</p>`);
});
await test('applying two sizes to the same word does not nest spans', async () => {
  await setDoc(`<p style="${P}">Hej fina världen</p><p style="${P}">Slut</p>`);
  await select('fina');
  await page.selectOption('#sel-size', '20px');
  await page.selectOption('#sel-size', '24px');
  eq(await out(), `<p style="${P}">Hej <span style="font-size:24px;">fina</span> världen</p><p style="${P0}">Slut</p>`);
});
await test('bold via toolbar and Ctrl+B produces <b>', async () => {
  await setDoc(`<p style="${P}">Hej fina världen</p><p style="${P}">Slut</p>`);
  await select('fina');
  await page.click('[data-cmd="bold"]');
  eq(await out(), `<p style="${P}">Hej <b>fina</b> världen</p><p style="${P0}">Slut</p>`);
  await select('världen'); await key('Control+i');
  eq(await out(), `<p style="${P}">Hej <b>fina</b> <i>världen</i></p><p style="${P0}">Slut</p>`);
});
await test('block preset "Rubrik" turns the paragraph into a heading div', async () => {
  await setDoc(`<p style="${P}">Min rubrik</p><p style="${P}">Slut</p>`);
  await caret('Min rubrik', 3);
  await page.selectOption('#sel-block', 'h-sub');
  eq(await out(), `<div style="${H}">Min rubrik</div><p style="${P0}">Slut</p>`);
  eq(await page.$eval('#sel-block', s => s.value), 'h-sub', 'toolbar shows the current style');
});
await test('block preset applies to every selected paragraph (not the next one after triple-click)', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Två</p><p style="${P}">Tre</p>`);
  await page.evaluate(() => { const ps = document.querySelectorAll('#editor p'); const r = document.createRange(); r.setStart(ps[0].firstChild, 0); r.setStart(ps[0].firstChild, 0); r.setEnd(ps[1], 0); document.getElementById('editor').focus(); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await page.selectOption('#sel-block', 'h-sub');
  eq(await out(), `<div style="${H}">Ett</div><p style="${P}">Två</p><p style="${P0}">Tre</p>`);
});
await test('clear formatting resets inline formatting and block style', async () => {
  await setDoc(`<div style="${H}" data-preset="h-sub"><b>Rubrik</b> <span style="color:#c92a2a;">röd</span></div><p style="${P}">Slut</p>`);
  await caret('röd', 1);
  await page.click('[data-cmd="clear"]');
  eq(await out(), `<p style="${P}">Rubrik röd</p><p style="${P0}">Slut</p>`);
});
await test('clear formatting on a partial selection only unwraps the selected text', async () => {
  await setDoc(`<p style="${P}"><b>fet text här</b></p><p style="${P}">Slut</p>`);
  await select('text');
  await page.click('[data-cmd="clear"]');
  eq(await out(), `<p style="${P}"><b>fet </b>text<b> här</b></p><p style="${P0}">Slut</p>`);
});
await test('editing a preset updates text that uses it (and keeps manual overrides)', async () => {
  await setDoc(`<p style="${P}">A</p><p style="${P}">B</p><p style="${P}">Slut</p>`);
  await caret('A'); await page.selectOption('#sel-block', 'h-sub');
  await caret('B'); await page.selectOption('#sel-block', 'h-sub');
  await page.selectOption('#sel-mb', '20px');
  await page.click('.preset[data-id="h-sub"] .preset-edit');
  await page.fill('#se-css', 'margin: 0 0 6px 0;\nfont-size: 18px;\nfont-weight: 700;\ncolor: #c92a2a;');
  await page.click('.modal-foot .mbtn.primary');
  eq(await out(), `<div style="margin:0 0 6px 0;font-size:18px;font-weight:700;color:#c92a2a;">A</div><div style="margin:0 0 20px 0;font-size:18px;font-weight:700;color:#c92a2a;">B</div><p style="${P0}">Slut</p>`);
});
await test('a new style can be created from the current selection and is saved', async () => {
  await setDoc(`<p style="margin:0 0 30px 0;font-size:15px;">Special</p><p style="${P}">Slut</p>`);
  await caret('Special', 2);
  await page.click('#style-new');
  await page.fill('#se-name', 'Min stil');
  await page.click('#se-grab');
  await page.click('.modal-foot .mbtn.primary');
  const saved = await page.evaluate(() => JSON.parse(localStorage.getItem('mhe2.presets')).find(p => p.name === 'Min stil'));
  eq(saved && saved.css, 'margin:0 0 30px 0;font-size:15px;', 'saved css');
  eq(saved.type, 'block', 'type');
});

console.log('\nBoxes, lists and tables');
await test('selecting paragraphs and choosing a box wraps them', async () => {
  await setDoc(`<p style="${P}">Före</p><p style="${P}">Ett</p><p style="${P}">Två</p><p style="${P}">Efter</p>`);
  await select('Ett', 'Två');
  await page.selectOption('#sel-box', 'box-info');
  eq(await out(), `<p style="${P}">Före</p><div style="margin:0 0 14px 0;padding:12px 14px;background:#f7f9fc;border-left:4px solid var(--muBlue);border-radius:4px;"><p style="${P}">Ett</p><p style="${P0}">Två</p></div><p style="${P0}">Efter</p>`);
});
await test('Enter on an empty last line in a box leaves the box', async () => {
  await setDoc(`<div style="margin:0 0 14px 0;padding:12px;" data-preset="box-info"><p style="${P}">I rutan</p></div><p style="${P}">Slut</p>`);
  await caret('I rutan', 7);
  await key('Enter'); await key('Enter'); await type('Utanför');
  eq(await out(), `<div style="margin:0 0 14px 0;padding:12px;"><p style="${P0}">I rutan</p></div><p style="${P}">Utanför</p><p style="${P0}">Slut</p>`);
});
await test('bullet list toggle on and off', async () => {
  await setDoc(`<p style="${P}">A</p><p style="${P}">B</p><p style="${P}">Slut</p>`);
  await select('A', 'B');
  await page.click('[data-cmd="ul"]');
  eq(await out(), `<ul style="${UL}"><li style="${LI}">A</li><li style="${P0}">B</li></ul><p style="${P0}">Slut</p>`);
  await page.click('[data-cmd="ul"]');
  eq(await out(), `<p style="${P}">A</p><p style="${P}">B</p><p style="${P0}">Slut</p>`);
});
await test('Enter in a list adds an item; Enter on an empty item ends the list', async () => {
  await setDoc(`<ul style="${UL}"><li style="${LI}">Ett</li></ul><p style="${P}">Slut</p>`);
  await caret('Ett', 3);
  await key('Enter'); await type('Två'); await key('Enter'); await key('Enter'); await type('Efter');
  eq(await out(), `<ul style="${UL}"><li style="${LI}">Ett</li><li style="${P0}">Två</li></ul><p style="${P}">Efter</p><p style="${P0}">Slut</p>`);
});
await test('Tab indents a list item into a sub list, Shift+Tab outdents', async () => {
  await setDoc(`<ul style="${UL}"><li style="${LI}">Ett</li><li style="${LI}">Två</li></ul><p style="${P}">Slut</p>`);
  await caret('Två', 1);
  await key('Tab');
  eq(await out(), `<ul style="${UL}"><li style="${P0}">Ett<ul style="margin:6px 0 0 0;padding-left:20px;"><li style="${P0}">Två</li></ul></li></ul><p style="${P0}">Slut</p>`);
  await key('Shift+Tab');
  eq(await out(), `<ul style="${UL}"><li style="${LI}">Ett</li><li style="${P0}">Två</li></ul><p style="${P0}">Slut</p>`);
});
await test('insert table via dialog gives inline-styled table', async () => {
  await setDoc(`<p style="${P}">Före</p><p style="${P}"><br></p>`);
  await caretIn('p', 1);
  await page.click('[data-pop="insert"]');
  await page.click('.popover [data-k="table"]');
  await page.fill('#t-rows', '2'); await page.fill('#t-cols', '2');
  await page.click('.modal-foot .mbtn.primary');
  await settle();
  await type('Värde');
  const html = await out();
  eq(html, `<p style="${P}">Före</p><table style="border-collapse:collapse;width:100%;margin:0;"><tr><th style="border:1px solid #d0d7de;padding:6px 8px;background:#f3f6fa;font-weight:600;text-align:left;">Rubrik 1</th><th style="border:1px solid #d0d7de;padding:6px 8px;background:#f3f6fa;font-weight:600;text-align:left;">Rubrik 2</th></tr><tr><td style="border:1px solid #d0d7de;padding:6px 8px;">Värde</td><td style="border:1px solid #d0d7de;padding:6px 8px;"></td></tr></table>`);
});
await test('Enter inside a table cell keeps the table intact', async () => {
  await setDoc(`<table style="border-collapse:collapse;"><tbody><tr><td style="padding:4px;">A1</td><td style="padding:4px;">B1</td></tr><tr><td style="padding:4px;">A2</td><td style="padding:4px;">B2</td></tr></tbody></table><p style="${P}">Slut</p>`);
  await caret('A1', 1);
  await key('Enter');
  eq(await out(), `<table style="border-collapse:collapse;"><tr><td style="padding:4px;">A<br>1</td><td style="padding:4px;">B1</td></tr><tr><td style="padding:4px;">A2</td><td style="padding:4px;">B2</td></tr></table><p style="${P0}">Slut</p>`);
});
await test('Tab moves between cells and adds a row after the last cell', async () => {
  await setDoc(`<table style="border-collapse:collapse;"><tbody><tr><td style="padding:4px;">A</td><td style="padding:4px;">B</td></tr></tbody></table><p style="${P}">Slut</p>`);
  await caret('A', 1);
  await key('Tab'); await type('b');
  await key('Tab'); await type('C');
  eq(await out(), `<table style="border-collapse:collapse;"><tr><td style="padding:4px;">A</td><td style="padding:4px;">b</td></tr><tr><td style="border:1px solid #d0d7de;padding:6px 8px;">C</td><td style="border:1px solid #d0d7de;padding:6px 8px;"></td></tr></table><p style="${P0}">Slut</p>`);
});
await test('table bar adds and removes columns', async () => {
  await setDoc(`<table style="border-collapse:collapse;"><tbody><tr><td style="padding:4px;">A</td></tr></tbody></table><p style="${P}">Slut</p>`);
  await caret('A', 1);
  await settle();
  await page.click('[data-tab="col-right"]');
  eq(await page.evaluate(() => document.querySelector('#editor tr').cells.length), 2, 'columns after add');
  await page.click('[data-tab="col-del"]');
  eq(await page.evaluate(() => document.querySelector('#editor tr').cells.length), 1, 'columns after delete');
});

console.log('\nDeleting and merging');
await test('Backspace at the start of a paragraph merges without junk spans', async () => {
  await setDoc(`<div style="${H}">Rubrik</div><p style="${P}">Text</p><p style="${P}">Slut</p>`);
  await caret('Text', 0);
  await key('Backspace');
  eq(await out(), `<div style="${H}">RubrikText</div><p style="${P0}">Slut</p>`);
  await type('!');
  eq(await out(), `<div style="${H}">Rubrik!Text</div><p style="${P0}">Slut</p>`);
});
await test('Backspace on an empty line removes it', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}"><br></p><p style="${P}">Två</p>`);
  await caretIn('p', 1);
  await key('Backspace');
  eq(await out(), `<p style="${P}">Ett</p><p style="${P0}">Två</p>`);
});
await test('Delete at the end of a paragraph pulls up the next one', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Två</p><p style="${P}">Slut</p>`);
  await caret('Ett', 3);
  await key('Delete');
  eq(await out(), `<p style="${P}">EttTvå</p><p style="${P0}">Slut</p>`);
});
await test('typing over a selection that spans two paragraphs merges them', async () => {
  await setDoc(`<p style="${P}">Första raden</p><p style="${P}">Andra raden</p><p style="${P}">Slut</p>`);
  await select('raden', 'Andra', 0, 5);
  await type('X');
  eq(await out(), `<p style="${P}">Första X raden</p><p style="${P0}">Slut</p>`);
});
await test('select all + Delete leaves one empty paragraph', async () => {
  await setDoc(`<div style="${H}">Rubrik</div><p style="${P}">Text</p><ul style="${UL}"><li style="${LI}">Punkt</li></ul>`);
  await caret('Text', 1);
  await key('Control+a'); await key('Delete');
  eq(await out(), '');
  await type('Nytt');
  eq(await out(), `<p style="${P0}">Nytt</p>`);
});

console.log('\nPaste and import');
await test('Word paste: mso junk removed, lists (incl. nested and numbered) rebuilt', async () => {
  await setDoc(`<p style="${P}"><br></p>`);
  await caretIn('p');
  await paste({ html: WORD_HTML, text: 'ignored' });
  const html = await out();
  for (const bad of ['mso', 'Mso', 'StartFragment', 'o:p', '·', 'Symbol', 'font-family', '\n']) ok(!html.includes(bad), `output contains "${bad}": ${html}`);
  eq(html, `<p style="${P}"><b>Instructions</b></p><p style="${P}">In this test, you will be shown one pair of words.</p><ul style="${UL}"><li style="${LI}">There may be several possible relationships.<ul style="margin:6px 0 0 0;padding-left:20px;"><li style="${P0}">Nested item</li></ul></li><li style="${P0}">Each task has only one correct answer.</li></ul><ol style="margin:0;padding-left:20px;"><li style="${LI}">First step</li><li style="${P0}">Second step</li></ol>`);
});
await test('Google Docs paste: keeps headings, bold, italic and lists; drops the fake <b> wrapper', async () => {
  await setDoc(`<p style="${P}"><br></p>`);
  await caretIn('p');
  await paste({ html: GDOCS_HTML, text: 'ignored' });
  eq(await out(), `<div style="${H}">Rubrik från Docs</div><p style="${P}"><b>Fet</b> och <i>kursiv</i></p><ul style="margin:0;padding-left:20px;"><li style="${LI}">Punkt ett</li><li style="${P0}">Punkt två</li></ul>`);
});
await test('plain text paste: lines become paragraphs, bullets become a list', async () => {
  await setDoc(`<p style="${P}"><br></p>`);
  await caretIn('p');
  await paste({ text: 'Inledning\n\n• Första\n• Andra\nAvslutning' });
  eq(await out(), `<p style="${P}">Inledning</p><ul style="${UL}"><li style="${LI}">Första</li><li style="${P0}">Andra</li></ul><p style="${P0}">Avslutning</p>`);
});
await test('pasting one line into the middle of a paragraph stays inline', async () => {
  await setDoc(`<p style="${P}">Hej världen</p><p style="${P}">Slut</p>`);
  await caret('Hej världen', 4);
  await paste({ html: '<p>fina </p>', text: 'fina ' });
  eq(await out(), `<p style="${P}">Hej fina världen</p><p style="${P0}">Slut</p>`);
});
await test('copy + paste inside the editor keeps styles', async () => {
  await setDoc(`<div style="${H}" data-preset="h-sub">Rubrik</div><p style="${P}">Text</p><p style="${P}"><br></p>`);
  await select('Rubrik', 'Text');
  const clip = await page.evaluate(() => { const dt = new DataTransfer(); document.getElementById('editor').dispatchEvent(new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true })); return dt.getData('text/html'); });
  await caretIn('p', 1);
  await paste({ html: clip, text: 'Rubrik\nText' });
  eq(await out(), `<div style="${H}">Rubrik</div><p style="${P}">Text</p><div style="${H}">Rubrik</div><p style="${P0}">Text</p>`);
});

await test('a copied heading line pasted into an empty line stays a heading', async () => {
  await setDoc(`<div style="${H}" data-preset="h-sub">Rubrik</div><p style="${P}">Text</p><p style="${P}"><br></p>`);
  await page.evaluate(() => { const ed = document.getElementById('editor'); const r = document.createRange(); r.setStart(ed.firstChild.firstChild, 0); r.setEnd(ed.children[1], 0); ed.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); });
  const clip = await page.evaluate(() => { const dt = new DataTransfer(); document.getElementById('editor').dispatchEvent(new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true })); return dt.getData('text/html'); });
  await caretIn('p', 1);
  await paste({ html: clip, text: 'Rubrik' });
  eq(await out(), `<div style="${H}">Rubrik</div><p style="${P}">Text</p><div style="margin:0;font-size:16px;font-weight:600;color:var(--muBlue);">Rubrik</div>`);
});
await test('part of a heading pasted into a paragraph joins the paragraph (like Word)', async () => {
  await setDoc(`<div style="${H}">Rubrik</div><p style="${P}">Text </p><p style="${P}">Slut</p>`);
  await select('Rubrik');
  const clip = await page.evaluate(() => { const dt = new DataTransfer(); document.getElementById('editor').dispatchEvent(new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true })); return dt.getData('text/html'); });
  await caret('Text', 5);
  await paste({ html: clip, text: 'Rubrik' });
  eq(await out(), `<div style="${H}">Rubrik</div><p style="${P}">Text Rubrik</p><p style="${P0}">Slut</p>`);
});
await test('web text pasted into an empty heading line takes the heading style', async () => {
  await setDoc(`<div style="${H}" data-preset="h-sub"><br></div><p style="${P}">Slut</p>`);
  await caretIn('div');
  await paste({ html: '<p style="font-size:30px">Från webben</p>', text: 'Från webben' });
  eq(await out(), `<div style="${H}">Från webben</div><p style="${P0}">Slut</p>`);
});
await test('multi-line paste into a list item creates list items', async () => {
  await setDoc(`<ul style="${UL}"><li style="${LI}">Ett</li></ul><p style="${P}">Slut</p>`);
  await caret('Ett', 3);
  await paste({ text: '\nTvå\nTre' });
  eq(await out(), `<ul style="${UL}"><li style="${LI}">Ett</li><li style="${LI}">Två</li><li style="${P0}">Tre</li></ul><p style="${P0}">Slut</p>`);
});
await test('paragraphs pasted into a table cell become line breaks', async () => {
  await setDoc(`<table style="border-collapse:collapse;"><tbody><tr><td style="padding:4px;"><br></td></tr></tbody></table><p style="${P}">Slut</p>`);
  await caretIn('td');
  await paste({ html: '<p>Rad 1</p><p>Rad 2</p>', text: 'Rad 1\nRad 2' });
  eq(await out(), `<table style="border-collapse:collapse;"><tr><td style="padding:4px;">Rad 1<br>Rad 2</td></tr></table><p style="${P0}">Slut</p>`);
});
await test('Word table paste keeps the table and applies default cell styles', async () => {
  await setDoc(`<p style="${P}"><br></p>`);
  await caretIn('p');
  await paste({ html: '<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style="border-collapse:collapse;border:none;mso-yfti-tbllook:1184"><tr><td width=302 valign=top style="width:226.55pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt"><p class=MsoNormal><b>Namn<o:p></o:p></b></p></td><td style="border:solid windowtext 1.0pt"><p class=MsoNormal>Värde<o:p></o:p></p></td></tr></table>', text: 'x' });
  eq(await out(), '<table style="border-collapse:collapse;width:100%;margin:0;"><tr><td style="border:1px solid #d0d7de;padding:6px 8px;"><b>Namn</b></td><td style="border:1px solid #d0d7de;padding:6px 8px;">Värde</td></tr></table>');
});
await test('"Ingen ruta" removes the box around the cursor', async () => {
  await setDoc(`<div style="padding:10px;" data-preset="box-info"><p style="${P}">Ett</p><p style="${P}">Två</p></div><p style="${P}">Slut</p>`);
  await caret('Två', 1);
  await page.selectOption('#sel-box', '');
  eq(await out(), `<p style="${P}">Ett</p><p style="${P}">Två</p><p style="${P0}">Slut</p>`);
});
await test('a box around a heading and a whole list', async () => {
  await setDoc(`<div style="${H}">Rubrik</div><ul style="${UL}"><li style="${LI}">A</li><li style="${LI}">B</li></ul><p style="${P}">Slut</p>`);
  await select('Rubrik', 'A');
  await page.click('.preset[data-id="box-section"] .preset-apply');
  eq(await out(), `<div style="margin:0 0 14px 0;"><div style="${H}">Rubrik</div><ul style="margin:0;padding-left:20px;"><li style="${LI}">A</li><li style="${P0}">B</li></ul></div><p style="${P0}">Slut</p>`);
});
await test('Backspace at the start of the first list item turns it into a paragraph', async () => {
  await setDoc(`<ul style="${UL}"><li style="${LI}">Ett</li><li style="${LI}">Två</li></ul><p style="${P}">Slut</p>`);
  await caret('Ett', 0);
  await key('Backspace');
  eq(await out(), `<p style="${P}">Ett</p><ul style="${UL}"><li style="${P0}">Två</li></ul><p style="${P0}">Slut</p>`);
});
await test('Enter at the start of a paragraph adds an empty line above and keeps the cursor on the text', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Två</p>`);
  await caret('Två', 0);
  await key('Enter'); await type('X');
  eq(await out(), `<p style="${P}">Ett</p><p style="${P}">&nbsp;</p><p style="${P0}">XTvå</p>`);
});
await test('clearing formatting keeps links', async () => {
  await setDoc(`<p style="${P}"><span style="color:#c92a2a;">Läs <a href="https://example.com" style="color:#c92a2a;">mer</a></span></p><p style="${P}">Slut</p>`);
  await select('Läs', 'mer');
  await page.click('[data-cmd="clear"]');
  eq(await out(), `<p style="${P}">Läs <a href="https://example.com" style="color:var(--muBlue);">mer</a></p><p style="${P0}">Slut</p>`);
});
await test('a block style from the sidebar applies to the paragraph at the cursor', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Slut</p>`);
  await caret('Ett', 1);
  await page.click('.preset[data-id="h-main"] .preset-apply');
  eq(await out(), `<div style="margin:0 0 10px 0;font-size:20px;font-weight:700;color:var(--muBlue);">Ett</div><p style="${P0}">Slut</p>`);
});

console.log('\nUndo, persistence, v1 migration');
await test('undo and redo restore the document', async () => {
  await setDoc(`<p style="${P}">Ett</p><p style="${P}">Slut</p>`);
  await caret('Ett', 1);
  await page.selectOption('#sel-block', 'h-sub');
  const after = await out();
  await caret('Ett', 1);
  await key('Control+z');
  eq(await out(), `<p style="${P}">Ett</p><p style="${P0}">Slut</p>`, 'after undo');
  await key('Control+y');
  eq(await out(), after, 'after redo');
});
await test('content survives a reload', async () => {
  await setDoc(`<p style="${P}">Sparad</p>`);
  await caret('Sparad', 6); await type(' text');
  await page.waitForTimeout(600);
  await page.reload(); await page.waitForFunction(() => window.MHE);
  eq(await out(), `<p style="${P0}">Sparad text</p>`);
});
await test('v1 data (styles, wrapper, content) is migrated on first start', async () => {
  await page.evaluate(() => {
    localStorage.clear();
    localStorage.setItem('mhe_styles', JSON.stringify([{ name: 'Brödtext', css: 'margin:0 0 12px 0;' }, { name: 'Grön text', css: 'color:#2b8a3e;' }, { name: 'Min box', css: 'padding:10px;background:#eee;' }]));
    localStorage.setItem('mhe_wrap', 'max-width:600px;font-family:Georgia,serif;');
    localStorage.setItem('mhe_content', '<p style="margin:0 0 12px 0;">Gammalt innehåll</p>');
  });
  await page.goto(APP); await page.waitForFunction(() => window.MHE);
  const names = await page.evaluate(() => MHE.presets.map(p => p.name + ':' + p.type));
  ok(names.includes('Grön text:text') && names.includes('Min box:box'), 'migrated presets: ' + names.join(', '));
  eq(names.filter(n => n.startsWith('Brödtext:')).length, 1, 'no duplicate default');
  eq(await out({}), '<div style="max-width:600px;font-family:Georgia,serif;"><p style="margin:0;">Gammalt innehåll</p></div>');
});

ok(true);
await browser.close();
const failed = results.filter(r => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} tester gick igenom.`);
if (pageErrors.length) { console.log('Sidfel:', pageErrors); }
process.exit(failed.length || pageErrors.length ? 1 : 0);
