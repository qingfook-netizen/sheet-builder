// Decision tree drawn as a picture (squares = decisions, circles = chance, red = best path) and pasted
// onto its own sheet. Numbers are read from the model when the tree is drawn.
/* global sbwFixRefs, sbwSheetName, sbwSetFormulas */

const SBT = { maxLeaves: 48, maxDepth: 8, colW: 230, leafGap: 66, top: 70, left: 30, boxW: 112, boxH: 54, circR: 32, circX: 12 };

function sbtLayout(root) {
  const nodes = [];
  let leafIdx = 0;
  const walk = (n, depth, parent, branch) => {
    const kids = (n.branches || []).filter((b) => b && b.node);
    const me = { n, depth: Math.min(depth, SBT.maxDepth), parent, branch, kids: [] };
    nodes.push(me);
    if (!kids.length || depth >= SBT.maxDepth || leafIdx >= SBT.maxLeaves) { me.leaf = true; me.slot = leafIdx++; }
    else {
      for (const b of kids) me.kids.push(walk(b.node, depth + 1, me, b));
      me.slot = (me.kids[0].slot + me.kids[me.kids.length - 1].slot) / 2;
    }
    return me;
  };
  walk(root, 0, null, null);
  return { nodes, leaves: leafIdx, maxDepth: Math.max(...nodes.map((x) => x.depth)) };
}

function sbtFmt(v, kind) {
  if (typeof v === "number" && isFinite(v)) {
    if (kind === "prob") return v.toLocaleString("en-US", { maximumFractionDigits: 4 });
    const whole = Math.abs(v - Math.round(v)) < 1e-9;
    return v.toLocaleString("en-US", { minimumFractionDigits: whole ? 0 : 2, maximumFractionDigits: 2 });
  }
  return v == null ? "" : String(v);
}

// draws onto a canvas; values: Map(node -> {value, prob}) already resolved to numbers/text
function sbtRender(canvas, tree, layout, vals) {
  const { nodes, leaves, maxDepth } = layout;
  const RED = "#C0262D", INK = "#262626";
  const W = SBT.left + maxDepth * SBT.colW + 230, H = SBT.top + Math.max(leaves - 1, 0) * SBT.leafGap + 70;
  const S = 2;
  canvas.width = W * S; canvas.height = H * S;
  const g = canvas.getContext("2d");
  g.scale(S, S);
  g.fillStyle = "#fff"; g.fillRect(0, 0, W, H);
  const font = (px, bold) => `${bold ? "600 " : ""}${px}px -apple-system, "Helvetica Neue", Arial, sans-serif`;
  const X = (x) => SBT.left + x.depth * SBT.colW + SBT.boxW / 2;
  const Y = (x) => SBT.top + x.slot * SBT.leafGap;
  const kind = (x) => x.leaf ? "end" : (String(x.n.type).toLowerCase() === "decision" ? "decision" : "chance");
  const halfW = (x) => kind(x) === "decision" ? SBT.boxW / 2 : SBT.circR + SBT.circX;
  const onBest = (x) => !!(x.branch && x.branch.best);

  // branches first, so nodes sit on top of the line ends
  for (const x of nodes) {
    if (!x.parent) continue;
    const p = x.parent, best = onBest(x);
    const x1 = X(p) + halfW(p), y1 = Y(p);
    const x2 = x.leaf ? X(x) - SBT.boxW / 2 - 8 : X(x) - halfW(x) - 2, y2 = Y(x);
    g.strokeStyle = best ? RED : INK; g.fillStyle = best ? RED : INK; g.lineWidth = best ? 2.4 : 1.3;
    g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke();
    const a = Math.atan2(y2 - y1, x2 - x1), L = 11;
    g.beginPath(); g.moveTo(x2, y2);
    g.lineTo(x2 - L * Math.cos(a - 0.38), y2 - L * Math.sin(a - 0.38));
    g.lineTo(x2 - L * Math.cos(a + 0.38), y2 - L * Math.sin(a + 0.38)); g.closePath(); g.fill();
    // label: "High 0.55" / "sell 80" beside the middle of the arrow, on the side away from the line
    const v = vals.get(x) || {};
    const parts = String(x.branch.label || "").split(/,\s*/);
    const l1 = [parts[0], v.prob !== undefined && v.prob !== "" ? sbtFmt(v.prob, "prob") : ""].filter(Boolean).join(" ");
    const lines = [l1, ...parts.slice(1)].filter(Boolean);
    const mx = (x1 + x2) / 2, my = (y1 + y2) / 2, lh = 16;
    g.font = font(13, best && kind(p) === "decision"); g.fillStyle = best ? RED : INK;
    if (Math.abs(y2 - y1) < 4) {
      g.textAlign = "center"; lines.forEach((t, i) => g.fillText(t, mx, my - 8 - (lines.length - 1 - i) * lh));
    } else if (y2 < y1) {                       // rising arrow: text above the right half, clear of the line
      const w = Math.max(...lines.map((t) => g.measureText(t).width));
      const tx = mx - 4, yAt = (xx) => y1 + (y2 - y1) * (xx - x1) / (x2 - x1);
      const base = Math.min(yAt(tx + w), y1) - 7;
      g.textAlign = "left"; lines.forEach((t, i) => g.fillText(t, tx, base - (lines.length - 1 - i) * lh));
    } else {                                    // falling arrow: text below-left
      g.textAlign = "right"; lines.forEach((t, i) => g.fillText(t, mx - 6, my + 18 + i * lh));
    }
  }
  // nodes
  for (const x of nodes) {
    const cx = X(x), cy = Y(x), v = vals.get(x) || {};
    const red = onBest(x) && !x.leaf;
    if (x.leaf) {
      g.font = font(14, true); g.fillStyle = INK; g.textAlign = "left"; g.textBaseline = "middle";
      const lab = x.n.label ? x.n.label + " = " : "";
      g.fillText(lab + sbtFmt(v.value), cx - SBT.boxW / 2, cy);
      g.textBaseline = "alphabetic";
      continue;
    }
    g.lineWidth = red ? 2.6 : 1.5; g.strokeStyle = red ? RED : INK; g.fillStyle = "#fff";
    g.beginPath();
    if (kind(x) === "decision") g.rect(cx - SBT.boxW / 2, cy - SBT.boxH / 2, SBT.boxW, SBT.boxH);
    else g.ellipse(cx, cy, SBT.circR + SBT.circX, SBT.circR, 0, 0, Math.PI * 2);
    g.fill(); g.stroke();
    // label inside, wrapped to two lines
    g.font = font(12.5); g.fillStyle = INK; g.textAlign = "center";
    const words = String(x.n.label || "").split(/\s+/), max = kind(x) === "decision" ? SBT.boxW - 10 : 2 * SBT.circR;
    const out = [""];
    for (const w of words) { const t = (out[out.length - 1] + " " + w).trim(); if (g.measureText(t).width > max && out[out.length - 1]) out.push(w); else out[out.length - 1] = t; }
    const ls = out.slice(0, 3);
    ls.forEach((t, i) => g.fillText(t, cx, cy + 4 + (i - (ls.length - 1) / 2) * 15));
    if (v.value !== undefined && v.value !== "") {
      g.font = font(13.5, true); g.fillStyle = INK;
      g.fillText("EMV " + sbtFmt(v.value), cx, cy - (kind(x) === "decision" ? SBT.boxH / 2 : SBT.circR) - 10);
    }
  }
  return { cssW: W, cssH: H };
}

async function drawTree(context, tree, map, used) {
  const fails = [];
  if (!tree || !tree.root) return { fails: ["Decision tree skipped: no root node."] };
  const sheetName = sbwSheetName(tree.sheet || "Decision Tree", used);
  const ws = context.workbook.worksheets.add(sheetName);
  const layout = sbtLayout(tree.root);
  ws.getRange("A1").values = [[String(tree.title || "Decision tree")]];
  ws.getRange("A1").format.font.bold = true; ws.getRange("A1").format.font.size = 14;
  ws.getRange("A2").values = [["Square = decision, circle = chance, red = best path. Numbers are taken from the model when the tree was drawn; the model sheet has the live values."]];
  ws.getRange("A2").format.font.italic = true; ws.getRange("A2").format.font.color = "#595959";
  ws.showGridlines = false;

  // evaluate every value/probability through Excel itself, in scratch cells, then clear them
  const asks = [];
  for (const x of layout.nodes) {
    if (x.n.value !== undefined && x.n.value !== null && x.n.value !== "") asks.push({ x, k: "value", e: x.n.value });
    if (x.branch && x.branch.prob !== undefined && x.branch.prob !== null && x.branch.prob !== "") asks.push({ x, k: "prob", e: x.branch.prob });
  }
  const SCR = 200;                                       // column GR: far off to the right
  asks.forEach((a, i) => {
    const raw = typeof a.e === "number" ? String(a.e) : String(a.e).trim().replace(/^=/, "");
    a.f = "=" + sbwFixRefs(raw, map);
    a.range = () => ws.getRangeByIndexes(i, SCR, 1, 1);
  });
  await sbwSetFormulas(context, asks.map((a) => ({ range: a.range, f: a.f, text: "?" })));
  const loaded = asks.map((a) => { const r = a.range(); r.load("values"); return r; });
  await context.sync();
  const vals = new Map();
  asks.forEach((a, i) => {
    const v = loaded[i].values[0][0];
    const o = vals.get(a.x) || {}; o[a.k] = v; vals.set(a.x, o);
    if (typeof v === "string" && /^#/.test(v)) fails.push(`${sheetName}: "${a.e}" gave ${v}`);
  });
  if (asks.length) ws.getRangeByIndexes(0, SCR, asks.length, 1).clear();

  // draw and paste the picture
  const canvas = document.createElement("canvas");
  const { cssW } = sbtRender(canvas, tree, layout, vals);
  const b64 = canvas.toDataURL("image/png").split(",")[1];
  const anchor = ws.getRange("A4");
  anchor.load("left,top");
  await context.sync();
  const img = ws.shapes.addImage(b64);
  img.lockAspectRatio = true;
  img.left = anchor.left; img.top = anchor.top;
  img.width = cssW * 0.75;                              // CSS px -> points, drawn at 2x for sharpness
  img.name = "Decision tree";
  try { await context.sync(); }
  catch (e) { fails.push(`${sheetName}: the tree picture could not be inserted (${e.message})`); }
  return { name: sheetName, fails };
}

if (typeof module !== "undefined") module.exports = { drawTree, sbtLayout, sbtRender };
