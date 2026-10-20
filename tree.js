// Draws a decision tree on its own sheet: rectangles = decisions, ellipses = chance nodes, arrows = branches.
// EMVs, probabilities and payoffs are live formulas in cells next to the shapes, linked to the calc sheet.
/* global sbwFixRefs, sbwSheetName */

const SBT = { leafGap: 4, firstRow: 6, nodeW: 76, nodeH: 42, nodeCol: 84, gapCol: 54, maxLeaves: 48, maxDepth: 8 };

// "=Calc!B3", "Calc!B3", 0.55, "0.55" -> Excel expression (no leading =)
function sbtExpr(v, map) {
  if (v === undefined || v === null || v === "") return null;
  if (typeof v === "number") return String(v);
  const s = String(v).trim().replace(/^=/, "");
  if (/^-?\d+(\.\d+)?%?$/.test(s)) return s;
  return "(" + sbwFixRefs(s, map) + ")";
}
const sbtStr = (s) => '"' + String(s).replace(/"/g, '""') + '"';

// lay out: leaves get rows top to bottom; a parent sits at the midpoint of its first and last child
function sbtLayout(root) {
  const nodes = [];
  let nextRow = SBT.firstRow, leaves = 0;
  const walk = (n, depth, parent, branch) => {
    const kids = (n.branches || []).filter((b) => b && b.node);
    const me = { n, depth: Math.min(depth, SBT.maxDepth), parent, branch, kids: [] };
    nodes.push(me);
    if (!kids.length || depth >= SBT.maxDepth || leaves >= SBT.maxLeaves) {
      me.leaf = true; me.row = nextRow; nextRow += SBT.leafGap; leaves++;
    } else {
      for (const b of kids) me.kids.push(walk(b.node, depth + 1, me, b));
      me.row = Math.round((me.kids[0].row + me.kids[me.kids.length - 1].row) / 2);
    }
    return me;
  };
  walk(root, 0, null, null);
  return { nodes, lastRow: nextRow, maxDepth: Math.max(...nodes.map((x) => x.depth)) };
}

const sbtNodeCol = (d) => 2 + d * 3;            // B, E, H, K ... (1-based)

async function drawTree(context, tree, map, used) {
  const fails = [];
  if (!tree || !tree.root) return { fails: ["Decision tree skipped: no root node."] };
  const sheetName = sbwSheetName(tree.sheet || "Decision Tree", used);
  const ws = context.workbook.worksheets.add(sheetName);
  const { nodes, lastRow, maxDepth } = sbtLayout(tree.root);
  const fmtV = String(tree.numFmt || "#,##0.##");

  // grid: fixed column widths so shapes line up with the label cells
  ws.getRange("A:A").format.columnWidth = 14;
  for (let d = 0; d <= maxDepth; d++) {
    const c = sbtNodeCol(d);
    ws.getRangeByIndexes(0, c - 1, 1, 1).getEntireColumn().format.columnWidth = SBT.nodeCol;
    ws.getRangeByIndexes(0, c, 1, 2).getEntireColumn().format.columnWidth = SBT.gapCol;
  }
  ws.getRange("A1").values = [[String(tree.title || "Decision tree")]];
  ws.getRange("A1").format.font.bold = true; ws.getRange("A1").format.font.size = 14;
  ws.getRange("A2").values = [["Rectangle = decision, ellipse = chance. EMVs, probabilities and payoffs are linked to the calculations. Red = best choice when the sheet was built."]];
  ws.getRange("A2").format.font.italic = true; ws.getRange("A2").format.font.color = "#595959";
  ws.showGridlines = false;

  // load real cell positions (Excel rounds widths) for every row/column we draw on
  const colPos = {}, rowPos = {};
  for (let d = 0; d <= maxDepth; d++) for (const c of [sbtNodeCol(d), sbtNodeCol(d) + 1]) { colPos[c] = ws.getRangeByIndexes(0, c - 1, 1, 1); colPos[c].load("left,width"); }
  for (let r = 1; r <= lastRow; r++) { rowPos[r] = ws.getRangeByIndexes(r - 1, 0, 1, 1); rowPos[r].load("top,height"); }
  await context.sync();
  const cx = (c) => colPos[c].left + colPos[c].width / 2;
  const cy = (r) => rowPos[r].top + rowPos[r].height / 2;
  const rowAt = (y) => { for (let r = 1; r <= lastRow; r++) if (y < rowPos[r].top + rowPos[r].height) return r; return lastRow; };

  // cells: EMV above each node, payoff at each leaf, branch label above each arrow
  const formulas = [];
  const put = (r, c, f, opts = {}) => formulas.push({ r, c, f, ...opts });
  for (const x of nodes) {
    const c = sbtNodeCol(x.depth);
    const val = sbtExpr(x.n.value, map);
    if (x.leaf) {
      const label = x.n.label ? x.n.label + " = " : "";
      if (val) put(x.row, c, `=${sbtStr(label)}&TEXT(${val},${sbtStr(fmtV)})`, { bold: true });
      else put(x.row, c, null, { text: x.n.label || "" , bold: true });
    } else if (val) {
      put(x.row - 2, c, `=${sbtStr("EMV ")}&TEXT(${val},${sbtStr(fmtV)})`, { bold: true, center: true });
    }
    if (x.branch) {
      const p = sbtExpr(x.branch.prob, map);
      const lbl = String(x.branch.label || "");
      const gc = sbtNodeCol(x.parent.depth) + 1;
      // arrow end points (same maths as the shapes below)
      const x1 = cx(gc - 1) + SBT.nodeW / 2, y1 = cy(x.parent.row);
      const x2 = x.leaf ? colPos[sbtNodeCol(x.depth)].left + 2 : cx(sbtNodeCol(x.depth)) - SBT.nodeW / 2, y2 = cy(x.row);
      const yAt = (xx) => y1 + (y2 - y1) * Math.min(1, Math.max(0, (xx - x1) / Math.max(1, x2 - x1)));
      const xm = colPos[gc].left + colPos[gc].width;            // label is centred on the two gap columns
      const hw = (lbl.length + (p ? 5 : 0)) * 2.9 + 4;
      const ya = yAt(xm - hw), yb = yAt(xm + hw);
      // upward or flat arrow: label above the line; downward arrow: label below it
      const r = y2 <= y1 ? rowAt(Math.min(ya, yb)) - 1 : rowAt(Math.max(ya, yb)) + 1;
      const f = p ? `=${sbtStr(lbl + (lbl ? " " : ""))}&TEXT(${p},"0.0##")` : null;
      put(Math.max(r, 3), gc, f, { text: lbl, red: !!x.branch.best, span: 2 });
    }
  }
  const cellOf = (x) => ws.getRangeByIndexes(x.r - 1, x.c - 1, 1, 1);
  const taken = new Set();
  for (const x of formulas) {
    const cell = cellOf(x);
    if (x.span) {                                  // branch label: merge the two gap cells and centre it
      const k = x.r + ":" + x.c;
      if (!taken.has(k)) { taken.add(k); ws.getRangeByIndexes(x.r - 1, x.c - 1, 1, x.span).merge(false); }
      cell.format.horizontalAlignment = "Center";
      cell.format.font.size = 10;
    }
    if (x.f) cell.formulas = [[x.f]]; else cell.values = [[x.text]];
    if (x.bold) cell.format.font.bold = true;
    if (x.red) cell.format.font.color = "#C00000";
    if (x.center) cell.format.horizontalAlignment = "Center";
  }
  try { await context.sync(); }
  catch {
    for (const x of formulas) {
      if (!x.f) continue;
      try { cellOf(x).formulas = [[x.f]]; await context.sync(); }
      catch { cellOf(x).values = [[x.text || "(check formula)"]]; await context.sync(); fails.push(`${sheetName}: could not link "${x.f}"`); }
    }
  }

  // shapes
  const red = new Set(nodes.filter((x) => x.branch && x.branch.best).map((x) => x));
  const box = (x) => {
    const c = sbtNodeCol(x.depth);
    return { l: cx(c) - SBT.nodeW / 2, t: cy(x.row) - SBT.nodeH / 2, r: cx(c) + SBT.nodeW / 2, mid: cy(x.row) };
  };
  for (const x of nodes) {
    if (x.leaf) continue;
    const b = box(x);
    const kind = String(x.n.type || "").toLowerCase() === "decision" ? "Rectangle" : "Ellipse";
    const s = ws.shapes.addGeometricShape(kind);
    s.left = b.l; s.top = b.t; s.width = SBT.nodeW; s.height = SBT.nodeH;
    s.fill.setSolidColor("#FFFFFF");
    s.lineFormat.color = red.has(x) ? "#C00000" : "#404040";
    s.lineFormat.weight = red.has(x) ? 2 : 1.25;
    s.textFrame.textRange.text = String(x.n.label || "").slice(0, 40);
    s.textFrame.textRange.font.size = 10;
    s.textFrame.textRange.font.color = "#1A1A1A";
    s.textFrame.horizontalAlignment = "Center";
    s.textFrame.verticalAlignment = "Middle";
    s.name = "SB tree node " + nodes.indexOf(x);
  }
  for (const x of nodes) {
    if (!x.parent) continue;
    const pb = box(x.parent);
    const c = sbtNodeCol(x.depth);
    const endX = x.leaf ? colPos[c].left + 2 : box(x).l;
    const ln = ws.shapes.addLine(pb.r, pb.mid, endX, cy(x.row), "Straight");
    const best = !!(x.branch && x.branch.best);
    ln.lineFormat.color = best ? "#C00000" : "#404040";
    ln.lineFormat.weight = best ? 2 : 1;
    ln.line.endArrowheadStyle = "Triangle";
  }
  try { await context.sync(); }
  catch (e) { fails.push(`${sheetName}: some tree shapes could not be drawn (${e.message})`); }
  return { name: sheetName, fails };
}

if (typeof module !== "undefined") module.exports = { drawTree, sbtLayout, sbtExpr };
