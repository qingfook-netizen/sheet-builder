// "Answers" sheet laid out like a model solution: Question heading, part labels, working lines,
// Excel function used, a highlighted final answer, and interpretation. Numbers are live links to the model.
/* global sbwFixRefs, sbwSheetName, sbwNum, sbwSetFormulas */

const SBA_STYLE = {
  working: { color: "#2E75B6" },
  excel:   { color: "#2E75B6", font: "Courier New" },
  answer:  { color: "#1A1A1A", bold: true, fill: "#E2EFDA" },
  note:    { color: "#2E75B6", italic: true }
};

// Excel string constants max out at 255 characters, so long text is joined in pieces
function sbaStr(s) {
  const out = [];
  for (let i = 0; i < s.length; i += 200) out.push('"' + s.slice(i, i + 200).replace(/"/g, '""') + '"');
  return out;
}

// "P(R | A) = 210 / 300 = {Calc!C7|0.0000}" -> ="P(R | A) = 210 / 300 = "&TEXT((Calc!C7),"0.0000")
function sbaLineFormula(text, map) {
  const parts = [], re = /\{([^{}|]+)(?:\|([^{}]+))?\}/g;
  let last = 0, m;
  const t = String(text || "");
  while ((m = re.exec(t))) {
    if (m.index > last) parts.push(...sbaStr(t.slice(last, m.index)));
    parts.push(sbwNum("(" + sbwFixRefs(m[1].trim().replace(/^=/, ""), map) + ")", m[2] && m[2].trim()));
    last = re.lastIndex;
  }
  if (last < t.length) parts.push(...sbaStr(t.slice(last)));
  return "=" + (parts.length ? parts.join("&") : '""');
}

async function writeAnswers(context, answers, map, used, position) {
  const ws = context.workbook.worksheets.add(sbwSheetName("Answers", used));
  ws.load("name");
  if (position != null) ws.position = position;
  ws.showGridlines = false;
  ws.getRange("A:A").format.columnWidth = 12;
  ws.getRange("B:B").format.columnWidth = 54;
  ws.getRange("C:C").format.columnWidth = 620;

  const items = [];
  let r = 1;
  for (const q of answers) {
    const head = ws.getRange("B" + r);
    head.values = [[[q.title, q.marks].filter(Boolean).join("   ")]];
    head.format.font.bold = true; head.format.font.size = 13;
    r += 2;
    for (const p of q.parts || []) {
      const lab = ws.getRange("B" + r);
      lab.values = [[String(p.label || "")]];
      lab.format.font.bold = true; lab.format.font.color = "#2E75B6";
      lab.format.verticalAlignment = "Top";
      for (const ln of p.lines || []) {
        const st = SBA_STYLE[ln.kind] || SBA_STYLE.working;
        const cell = ws.getRange("C" + r);
        const f = cell.format;
        f.font.color = st.color;
        if (st.bold) f.font.bold = true;
        if (st.italic) f.font.italic = true;
        if (st.font) f.font.name = st.font;
        if (st.fill) { f.fill.color = st.fill; ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach((e) => { const b = f.borders.getItem(e); b.style = "Continuous"; b.color = "#A9D08E"; }); }
        f.wrapText = true; f.verticalAlignment = "Top";
        const row = r;
        items.push({ range: () => ws.getRange("C" + row), f: sbaLineFormula(ln.text, map), text: String(ln.text || "") });
        r++;
      }
      r++;
    }
    r++;
  }
  await context.sync();
  const fails = await sbwSetFormulas(context, items);
  ws.getRange("B1:C" + Math.max(r, 1)).format.autofitRows();
  await context.sync();
  return { name: ws.name, fails: fails.map((t) => "Answers: could not link " + t) };
}

if (typeof module !== "undefined") module.exports = { writeAnswers, sbaLineFormula };
