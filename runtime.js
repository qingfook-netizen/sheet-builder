/* global Office, Excel, SB_PROMPT, writeSpec */
// Runs invisibly behind the ribbon buttons and the keyboard shortcut. No task pane.

const DEFAULTS = { apiKey: "", model: "anthropic/claude-opus-5.5", standing: "" };
const BASE = location.href.replace(/[^/]*([?#].*)?$/, "");
let busy = false;

function getSettings() {
  try { return { ...DEFAULTS, ...JSON.parse(localStorage.getItem("sb-settings") || "{}") }; }
  catch { return { ...DEFAULTS }; }
}
const saveSettings = (o) => localStorage.setItem("sb-settings", JSON.stringify({ ...getSettings(), ...o }));

// small popup for messages (Excel has no toast API for add-ins)
function popup(text) {
  Office.context.ui.displayDialogAsync(BASE + "message.html#" + encodeURIComponent(text), { height: 30, width: 30, displayInIframe: true }, (r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) return;
    const d = r.value;
    d.addEventHandler(Office.EventType.DialogMessageReceived, () => d.close());
  });
}

function parseSpec(text) {
  const t = String(text || "").trim().replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/, "");
  const a = t.indexOf("{"), b = t.lastIndexOf("}");
  if (a < 0 || b <= a) throw new Error("The reply contained no JSON.");
  return JSON.parse(t.slice(a, b + 1));
}

// the newest picture on the active sheet (the screenshot you just pasted)
async function grabPicture() {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("id,name");
    const shapes = ws.shapes;
    shapes.load("items/id,items/type");
    await ctx.sync();
    const pics = shapes.items.filter((s) => s.type === "Image");
    if (!pics.length) return null;
    const pic = pics[pics.length - 1];
    const img = pic.getAsImage(Excel.PictureFormat.png);
    await ctx.sync();
    return { base64: img.value, shapeId: pic.id, sheetId: ws.id };
  });
}

async function addPlaceholder() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const names = new Set(sheets.items.map((s) => s.name.toLowerCase()));
    let name = "Building", i = 2;
    while (names.has(name.toLowerCase())) name = "Building " + i++;
    const ws = sheets.add(name);
    ws.load("id");
    ws.getRange("A1").values = [["Building from screenshot… (sending)"]];
    ws.getRange("A1").format.font.color = "#595959";
    ws.activate();
    await ctx.sync();
    return ws.id;
  });
}

const status = (sheetId, text) => Excel.run(async (ctx) => {
  ctx.workbook.worksheets.getItem(sheetId).getRange("A1").values = [[text]];
  await ctx.sync();
}).catch(() => {});

async function callModel(base64, onProgress) {
  const o = getSettings();
  if (!o.apiKey) throw new Error("No API key yet. Click Settings on the Sheet Builder tab and paste your OpenRouter key.");
  const body = {
    model: o.model || DEFAULTS.model, stream: true, max_tokens: 16000,
    messages: [
      { role: "system", content: SB_PROMPT },
      { role: "user", content: [
        { type: "text", text: "Build the workbook these instructions describe." + (o.standing ? "\n\nStanding notes: " + o.standing : "") },
        { type: "image_url", image_url: { url: "data:image/png;base64," + base64 } }
      ] }
    ]
  };
  const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: { "Authorization": "Bearer " + o.apiKey, "Content-Type": "application/json", "X-Title": "Sheet Builder for Excel" },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const d = await res.json().catch(() => ({}));
    throw new Error(`OpenRouter ${res.status}: ${d?.error?.message || res.statusText}`);
  }
  const reader = res.body.getReader(), dec = new TextDecoder();
  let buf = "", out = "", last = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    const lines = buf.split("\n"); buf = lines.pop();
    for (const line of lines) {
      if (!line.startsWith("data: ") || line === "data: [DONE]") continue;
      let j; try { j = JSON.parse(line.slice(6)); } catch { continue; }
      if (j.error) throw new Error(j.error.message || "stream error");
      out += j.choices?.[0]?.delta?.content || "";
    }
    if (Date.now() - last > 1500) { last = Date.now(); onProgress(out.length); }
  }
  return out;
}

async function buildFromImage(event) {
  event?.completed?.();          // release the ribbon button; the work carries on in this runtime
  if (busy) return;
  busy = true;
  let sheetId = null;
  try {
    const pic = await grabPicture();
    if (!pic) {
      popup("No picture on this sheet. Take a screenshot to the clipboard (Cmd+Ctrl+Shift+4), paste it into the sheet with Cmd+V, then run Build again.");
      return;
    }
    sheetId = await addPlaceholder();
    const reply = await callModel(pic.base64, (n) => status(sheetId, `Building from screenshot… ${n.toLocaleString()} characters received`));
    await status(sheetId, "Writing the sheet…");
    const spec = parseSpec(reply);
    await Excel.run((ctx) => writeSpec(ctx, spec, sheetId));
    // remove the pasted screenshot now that it's been used
    await Excel.run(async (ctx) => { ctx.workbook.worksheets.getItem(pic.sheetId).shapes.getItem(pic.shapeId).delete(); await ctx.sync(); }).catch(() => {});
  } catch (e) {
    if (sheetId) await status(sheetId, "Error: " + e.message + "  (your screenshot is still on the original sheet; delete this sheet and try again)");
    else popup("Error: " + e.message);
  } finally {
    busy = false;
  }
}

function openSettings(event) {
  event?.completed?.();
  Office.context.ui.displayDialogAsync(BASE + "settings.html", { height: 55, width: 35, displayInIframe: true }, (r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) return;
    const d = r.value;
    d.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      try { const o = JSON.parse(arg.message); if (o && o.save) saveSettings(o.save); } catch {}
      d.close();
    });
  });
}

Office.onReady(() => {
  Office.actions.associate("buildFromImage", buildFromImage);
  Office.actions.associate("openSettings", openSettings);
});
