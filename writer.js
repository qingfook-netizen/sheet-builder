// Writes the model's JSON spec into the open workbook with the Office.js Excel API.
// Shared by runtime.js and the tests (which pass a mock context).

const SBW_PRESETS = {
  title:  { bold: true, size: 14 },
  header: { bold: true, fill: "#D9E1F2", border: "bottom", wrap: true },
  input:  { fontColor: "#0000FF", fill: "#FFF2CC" },     // blue on yellow = cells you change
  total:  { bold: true, border: "top" },
  note:   { italic: true, fontColor: "#595959", wrap: true }
};
const SBW_REF = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/;
const SBW_RANGE = /^\$?[A-Z]{1,3}\$?[1-9]\d{0,6}:\$?[A-Z]{1,3}\$?[1-9]\d{0,6}$/;

const sbwColNum = (L) => L.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const sbwColLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const sbwHex = (h) => { const s = String(h || "").trim(); return /^#?[0-9a-fA-F]{6}$/.test(s) ? "#" + s.replace(/^#/, "").toUpperCase() : null; };

function sbwSheetName(name, used) {
  let n = String(name || "Sheet").replace(/[\[\]:*?/\\]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31) || "Sheet";
  const base = n; let i = 2;
  while (used.has(n.toLowerCase())) n = (base.slice(0, 27) + " " + i++).slice(0, 31);
  used.add(n.toLowerCase());
  return n;
}

// point formulas at the renamed sheets ('Old name'!A1 or Old!A1 -> 'New name'!A1)
function sbwFixRefs(f, map) {
  if (!map.size) return f;
  const q = (n) => "'" + n.replace(/'/g, "''") + "'!";
  return f
    .replace(/'((?:[^']|'')+)'!/g, (m, n) => { const k = map.get(n.replace(/''/g, "'").toLowerCase()); return k ? q(k) : m; })
    .replace(/(^|[^A-Za-z0-9_.'!])([A-Za-z_][A-Za-z0-9_.]*)!/g, (m, pre, n) => { const k = map.get(n.toLowerCase()); return k ? pre + q(k) : m; });
}

function sbwStyle(range, c) {
  const p = SBW_PRESETS[c.style] || {};
  const pick = (k) => (c[k] != null ? c[k] : p[k]);
  const font = range.format.font;
  if (pick("bold") != null) font.bold = !!pick("bold");
  if (pick("italic") != null) font.italic = !!pick("italic");
  if (pick("size")) font.size = Number(pick("size"));
  const fc = sbwHex(pick("fontColor")); if (fc) font.color = fc;
  const fill = sbwHex(pick("fill")); if (fill) range.format.fill.color = fill;
  const b = pick("border");
  const edges = b === "all" || b === true ? ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] : b === "top" ? ["EdgeTop"] : b === "bottom" ? ["EdgeBottom"] : [];
  for (const e of edges) { const x = range.format.borders.getItem(e); x.style = "Continuous"; x.color = "#808080"; }
  if (pick("wrap")) { range.format.wrapText = true; range.format.verticalAlignment = "Top"; }
  const al = String(c.align || "").toLowerCase();
  if (["left", "center", "right"].includes(al)) range.format.horizontalAlignment = al[0].toUpperCase() + al.slice(1);
  if (c.numFmt) range.numberFormat = [[String(c.numFmt)]];
}

// context: Excel.RequestContext. placeholderId: the "Building…" sheet, reused as the first sheet.
async function writeSpec(context, spec, placeholderId) {
  if (!spec || !Array.isArray(spec.sheets) || !spec.sheets.length) throw new Error("The model returned no sheets.");
  const sheets = context.workbook.worksheets;
  sheets.load("items/name,items/id");
  await context.sync();
  const used = new Set(sheets.items.filter((s) => s.id !== placeholderId).map((s) => s.name.toLowerCase()));

  // 1) create / rename sheets
  const map = new Map(), made = [];
  spec.sheets.forEach((s, i) => {
    const name = sbwSheetName(s.name, used);
    if (String(s.name || "").toLowerCase() !== name.toLowerCase()) map.set(String(s.name || "").toLowerCase(), name);
    let ws;
    if (i === 0 && placeholderId) { ws = sheets.getItem(placeholderId); ws.getRange().clear(); ws.name = name; }
    else ws = sheets.add(name);
    made.push({ ws, name, s });
  });

  // 2) values, formats, widths, merges, dropdowns, freeze
  const formulas = [], skipped = [];
  let lastRow = 0;
  for (const { ws, s } of made) {
    for (const c of s.cells || []) {
      const m = SBW_REF.exec(String(c.ref || "").toUpperCase());
      if (!m) { skipped.push(String(c.ref)); continue; }
      const ref = m[1] + m[2];
      if (ws === made[0].ws) lastRow = Math.max(lastRow, Number(m[2]));
      const r = ws.getRange(ref);
      let f = c.formula;
      if (f == null && typeof c.value === "string" && c.value.startsWith("=") && c.value.length > 1) f = c.value;
      if (f != null && String(f).trim()) formulas.push({ ws, ref, f: "=" + sbwFixRefs(String(f).trim().replace(/^=/, ""), map) });
      else if (c.value !== undefined && c.value !== null && c.value !== "") r.values = [[c.value]];
      if (Array.isArray(c.list) && c.list.length) {
        r.dataValidation.clear();
        r.dataValidation.rule = { list: { inCellDropDown: true, source: c.list.map((x) => String(x).replace(/,/g, " ")).join(",") } };
      }
      sbwStyle(r, c);
    }
    for (const [col, w] of Object.entries(s.columnWidths || {})) {
      if (/^[A-Z]{1,3}$/.test(col) && Number(w) > 0) ws.getRange(col + ":" + col).format.columnWidth = Math.min(Number(w), 120) * 6;
    }
    for (const mr of s.merges || []) {
      const t = String(mr).toUpperCase();
      if (SBW_RANGE.test(t)) ws.getRange(t).merge(false); else skipped.push(t);
    }
    const fz = SBW_REF.exec(String(s.freeze || "").toUpperCase());
    if (fz) {
      const x = sbwColNum(fz[1]) - 1, y = Number(fz[2]) - 1;
      if (x > 0 && y > 0) ws.freezePanes.freezeAt(ws.getRange("A1:" + sbwColLetter(x) + y));
      else if (y > 0) ws.freezePanes.freezeRows(y);
      else if (x > 0) ws.freezePanes.freezeColumns(x);
    }
  }
  await context.sync();

  // 3) formulas: all at once; if Excel rejects one, retry one by one and keep the bad ones as text
  const bad = [];
  if (formulas.length) {
    for (const x of formulas) x.ws.getRange(x.ref).formulas = [[x.f]];
    try { await context.sync(); }
    catch {
      for (const x of formulas) {
        try { x.ws.getRange(x.ref).formulas = [[x.f]]; await context.sync(); }
        catch {
          const r = x.ws.getRange(x.ref);
          r.numberFormat = [["@"]]; r.values = [[x.f]]; r.format.font.color = "#C00000";
          await context.sync();
          bad.push(x.ref + ": " + x.f);
        }
      }
    }
  }

  // 4) notes under the first sheet's content
  const notes = [
    ...(spec.missing || []).map((t) => "Missing input: " + t),
    ...(spec.assumptions || []).map((t) => "Assumption: " + t),
    ...bad.map((t) => "Excel rejected this formula (left as red text): " + t),
    ...(skipped.length ? ["Skipped invalid references: " + skipped.join(", ")] : [])
  ];
  const first = made[0].ws;
  if (notes.length) {
    const r0 = lastRow + 2;
    const head = first.getRange("A" + r0);
    head.values = [["Sheet Builder notes (delete when done)"]];
    head.format.font.bold = true; head.format.font.color = "#595959";
    const body = first.getRange(`A${r0 + 1}:A${r0 + notes.length}`);
    body.values = notes.map((t) => [String(t)]);
    body.format.font.italic = true; body.format.font.color = "#595959";
  }
  first.activate();
  first.getRange("A1").select();
  await context.sync();
  return { sheets: made.map((x) => x.name), notes, bad, skipped };
}

if (typeof module !== "undefined") module.exports = { writeSpec, sbwFixRefs, sbwSheetName };
