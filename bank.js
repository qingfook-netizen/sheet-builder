// Instruction Bank: a sheet in the workbook that keeps what each build read, the open questions
// (with yellow cells for your answers), the original screenshots, and anything you paste in.
// Everything on it is sent with every build, and Rebuild works from it alone.
/* global Excel, sbwColLetter */

const SB_BANK = "Instruction Bank";

async function readBank(context, maxImages) {
  const ws = context.workbook.worksheets.getItemOrNullObject(SB_BANK);
  await context.sync();
  if (ws.isNullObject) return null;
  const used = ws.getUsedRangeOrNullObject(true);
  used.load("values,rowIndex,columnIndex");
  const shapes = ws.shapes;
  shapes.load("items/id,items/type,items/top");
  await context.sync();
  const lines = [];
  if (!used.isNullObject) {
    used.values.forEach((row, i) => {
      const cells = [];
      row.forEach((v, j) => { if (v !== "" && v !== null) cells.push(`${sbwColLetter(used.columnIndex + j + 1)}${used.rowIndex + i + 1}: ${v}`); });
      if (cells.length) lines.push(cells.join(" | "));
    });
  }
  let text = lines.join("\n");
  if (text.length > 60000) text = "[older entries trimmed]\n" + text.slice(-60000);
  const pics = shapes.items.filter((s) => s.type === "Image").sort((a, b) => a.top - b.top).slice(-maxImages);
  const imgs = pics.map((p) => p.getAsImage(Excel.PictureFormat.png));
  await context.sync();
  return { text, images: imgs.map((i) => i.value) };
}

async function writeBank(context, spec, picBase64, label) {
  let ws = context.workbook.worksheets.getItemOrNullObject(SB_BANK);
  await context.sync();
  if (ws.isNullObject) {
    ws = context.workbook.worksheets.add(SB_BANK);
    ws.getRange("A:A").format.columnWidth = 520;
    ws.getRange("B:B").format.columnWidth = 260;
    ws.getRange("C:C").format.columnWidth = 16;
    ws.getRange("A1").values = [["Instruction Bank"]];
    ws.getRange("A1").format.font.bold = true; ws.getRange("A1").format.font.size = 14;
    ws.getRange("A2").values = [["Everything on this sheet is sent with every build. Correct what was read, answer the open questions in the yellow cells, paste extra notes, data or screenshots anywhere, then click Rebuild from Bank."]];
    ws.getRange("A2").format.wrapText = true; ws.getRange("A2").format.font.color = "#595959";
  }
  const used = ws.getUsedRangeOrNullObject(false);
  used.load("rowIndex,rowCount");
  await context.sync();
  const start = used.isNullObject ? 4 : used.rowIndex + used.rowCount + 2;
  let r = start;
  const text = (cellRef, v) => { const c = ws.getRange(cellRef); c.numberFormat = [["@"]]; c.values = [[String(v)]]; c.format.wrapText = true; c.format.verticalAlignment = "Top"; return c; };
  const bold = (cellRef, v, fill) => { const c = text(cellRef, v); c.format.font.bold = true; if (fill) c.format.fill.color = fill; return c; };

  bold("A" + r, label, "#D9E1F2"); ws.getRange("B" + r).format.fill.color = "#D9E1F2"; r++;
  const instr = Array.isArray(spec.instructions) ? spec.instructions.join("\n\n") : String(spec.instructions || "");
  if (instr.trim()) {
    bold("A" + r++, "Instructions as read (edit if anything was misread)");
    for (const para of instr.split(/\n{2,}/).map((s) => s.trim()).filter(Boolean)) text("A" + r++, para.slice(0, 30000));
    r++;
  }
  const qs = (spec.questions || []).filter(Boolean);
  if (qs.length) {
    bold("A" + r, "Open questions (answer in column B, then Rebuild from Bank)"); bold("B" + r, "Your answer"); r++;
    for (const q of qs) {
      text("A" + r, q);
      const a = ws.getRange("B" + r); a.numberFormat = [["@"]]; a.format.fill.color = "#FFF2CC"; a.format.font.color = "#0000FF"; a.format.wrapText = true;
      r++;
    }
    r++;
  }
  const asm = (spec.assumptions || []).filter(Boolean);
  if (asm.length) {
    bold("A" + r++, "Assumptions used in this build");
    for (const t of asm) text("A" + r++, "• " + t);
    r++;
  }
  // keep the screenshot with the build that used it
  let endRow = r;
  if (picBase64) {
    const anchor = ws.getRange("D" + start);
    anchor.load("left,top");
    const img = ws.shapes.addImage(picBase64);
    img.load("width,height");
    await context.sync();
    img.lockAspectRatio = true;
    const w = Math.min(img.width, 460), h = img.height * (w / img.width);
    img.width = w; img.left = anchor.left; img.top = anchor.top;
    img.name = "SB screenshot " + label;
    endRow = Math.max(r, start + Math.ceil(h / 15) + 1);
  }
  const end = ws.getRange("A" + endRow);
  end.values = [["— end of " + label.split(" · ")[0].toLowerCase() + " —"]];
  end.format.font.color = "#A6A6A6"; end.format.font.size = 9;
  ws.getRange(`A${start}:B${endRow}`).format.autofitRows();
  await context.sync();
}

if (typeof module !== "undefined") module.exports = { readBank, writeBank, SB_BANK };
