// Microsoft Graph access to the workbook using the app's own credentials.
const G = "https://graph.microsoft.com/v1.0";
let tok = { v: null, exp: 0 }, itemRef = null, layout = null;

export async function token() {
  if (tok.v && Date.now() < tok.exp - 60000) return tok.v;
  const body = new URLSearchParams({ client_id: process.env.CLIENT_ID, client_secret: process.env.CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body });
  const j = await r.json();
  if (!r.ok) throw new Error("Token error: " + (j.error_description || r.status));
  tok = { v: j.access_token, exp: Date.now() + j.expires_in * 1000 };
  return j.access_token;
}
export async function graph(path, t, opts = {}) {
  const r = await fetch(G + path, { method: opts.method || "GET", body: opts.body,
    headers: { Authorization: "Bearer " + t, "Cache-Control": "no-cache", "Content-Type": "application/json", ...(opts.headers || {}) } });
  if (!r.ok) throw new Error(`Graph ${r.status} ${path.slice(0, 80)}: ${(await r.text()).slice(0, 300)}`);
  return r.status === 204 ? null : r.json();
}
const shareId = (u) => "u!" + Buffer.from(u).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
export const TABLE = () => encodeURIComponent(process.env.TABLE_NAME || "tbl_Dispatch");

export async function workbook(t) {
  if (!itemRef) {
    const it = await graph(`/shares/${shareId(process.env.SHARE_LINK)}/driveItem?$select=id,parentReference`, t);
    itemRef = { driveId: it.parentReference.driveId, itemId: it.id };
  }
  return `/drives/${itemRef.driveId}/items/${itemRef.itemId}/workbook`;
}
/** Opens a session and runs fn(wb, headers) inside it. persist=true for writes. */
export async function withSession(persist, fn) {
  const t = await token(); const wb = await workbook(t);
  const s = await graph(wb + "/createSession", t, { method: "POST", body: JSON.stringify({ persistChanges: persist }) });
  const h = { "workbook-session-id": s.id };
  try { return await fn({ t, wb, h, get: (p) => graph(wb + p, t, { headers: h }), patch: (p, body) => graph(wb + p, t, { method: "PATCH", headers: h, body: JSON.stringify(body) }) }); }
  finally { graph(wb + "/closeSession", t, { method: "POST", headers: h }).catch(() => {}); }
}
/** Table layout: sheet name, header row number, first column letter. Cached per instance. */
export async function tableLayout(s) {
  if (layout) return layout;
  const r = await s.get(`/tables/${TABLE()}/range?$select=address`);   // e.g. Dispatch_Log!A3:V503
  const m = /^(.*)!\$?([A-Z]+)\$?(\d+):/.exec(r.address);
  layout = { sheet: m[1].replace(/^'|'$/g, ""), col: m[2], headerRow: parseInt(m[3], 10) };
  return layout;
}
export async function readTable(s) {
  const [hdr, body] = await Promise.all([s.get(`/tables/${TABLE()}/headerRowRange?$select=values`), s.get(`/tables/${TABLE()}/rows?$select=values`)]);
  const headers = hdr.values[0].map((h) => String(h).trim());
  return { headers, rows: body.value.map((x, i) => ({ i, raw: x.values[0], obj: Object.fromEntries(headers.map((h, k) => [h, x.values[0][k]])) })) };
}
export const colLetter = (n) => { let s = ""; n += 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
export const serial = (iso) => (iso ? Math.round((Date.parse(String(iso).slice(0, 10) + "T00:00:00Z") / 86400000) + 25569) : "");
export const fromSerial = (v) => (typeof v === "number" && v > 20000) ? new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10) : (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v) ? v.slice(0, 10) : null);
