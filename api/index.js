// NMDC dispatch backend — single-file build (dispatch, sync, tform, arrive) — build 2026-09-21c

// lib/db.js
import { neon } from "@neondatabase/serverless";
var sql = neon(process.env.DATABASE_URL);
var STATUSES = ["Under Loading", "Loaded", "In-Transit", "Arrived"];
function json(res, status, body) {
  res.setHeader("Cache-Control", "no-store");
  res.status(status).json(body);
}
var d = (v) => v ? String(v).slice(0, 10) : null;
function cors(req, res) {
  const norm = (u) => String(u || "").trim().toLowerCase().replace(/\/+$/, "");
  const allowed = (process.env.ALLOWED_ORIGINS || "").split(",").map(norm).filter(Boolean);
  const origin = req.headers.origin;
  const ok = origin && (allowed.includes("*") || allowed.some((a) => a === norm(origin) || a === norm(origin).replace(/^https?:\/\//, "")));
  if (ok) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-dash-pin");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Max-Age", "86400");
  }
  if (req.method === "OPTIONS") {
    res.status(204).end();
    return true;
  }
  return false;
}

// api/dispatch.js
async function handler(req, res) {
  if (cors(req, res)) return;
  if (!process.env.DASH_PIN) return json(res, 503, { error: "DASH_PIN not set" });
  if (req.headers["x-dash-pin"] !== process.env.DASH_PIN) return json(res, 401, { error: "pin" });
  try {
    const c = await sql`select * from coils order by entry_date nulls first, created_at`;
    const rows = c.map((r) => ({
      "Entry Date": r.entry_date,
      "NMDC SO No": r.so_no,
      "NMDC Invoice / DO No": r.do_no,
      "Coil No": r.coil_no,
      "Form": r.form,
      "Grade": r.grade,
      "Thk (mm)": r.thk == null ? "" : Number(r.thk),
      "Width (mm)": r.width ?? "",
      "Length (mm)": r.length ?? "",
      "Weight (MT)": r.weight == null ? "" : Number(r.weight),
      "Status": r.status,
      "Vehicle No": r.vehicle_no || "",
      "Transporter Name": r.transporter_name || "",
      "Transporter Mobile": r.transporter_mobile || "",
      "LR No": r.lr_no || "",
      "Dispatch Date (ex NMDC)": r.dispatch_date,
      "Expected Arrival": r.expected_arrival,
      "Actual Arrival": r.actual_arrival,
      "Destination": r.destination || "",
      "Remarks": r.remarks || "",
      "Entered By": r.entered_by || "BSC"
    }));
    json(res, 200, { rows, fetchedAt: (/* @__PURE__ */ new Date()).toISOString() });
  } catch (e) {
    json(res, 502, { error: e.message });
  }
}

// api/sync.js
import crypto from "node:crypto";

// lib/graph.js
var G = "https://graph.microsoft.com/v1.0";
var tok = { v: null, exp: 0 };
var itemRef = null;
var layout = null;
async function token() {
  if (tok.v && Date.now() < tok.exp - 6e4) return tok.v;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });
  const r = await fetch(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );
  const j = await r.json();
  if (!r.ok) throw new Error("Token error: " + (j.error_description || r.status));
  tok = { v: j.access_token, exp: Date.now() + j.expires_in * 1e3 };
  return j.access_token;
}
async function graph(path, t, opts = {}) {
  const r = await fetch(G + path, {
    method: opts.method || "GET",
    body: opts.body,
    headers: { Authorization: "Bearer " + t, "Cache-Control": "no-cache", "Content-Type": "application/json", ...opts.headers || {} }
  });
  if (!r.ok) throw new Error(`Graph ${r.status} ${path.slice(0, 80)}: ${(await r.text()).slice(0, 300)}`);
  return r.status === 204 ? null : r.json();
}
var shareId = (u) => "u!" + Buffer.from(u).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
var TABLE = () => encodeURIComponent(process.env.TABLE_NAME || "tbl_Dispatch");
async function workbook(t) {
  if (!itemRef) {
    const it = await graph(`/shares/${shareId(process.env.SHARE_LINK)}/driveItem?$select=id,parentReference`, t);
    itemRef = { driveId: it.parentReference.driveId, itemId: it.id };
  }
  return `/drives/${itemRef.driveId}/items/${itemRef.itemId}/workbook`;
}
async function withSession(persist, fn) {
  const t = await token();
  const wb = await workbook(t);
  const s2 = await graph(wb + "/createSession", t, { method: "POST", body: JSON.stringify({ persistChanges: persist }) });
  const h = { "workbook-session-id": s2.id };
  try {
    return await fn({ t, wb, h, get: (p) => graph(wb + p, t, { headers: h }), patch: (p, body) => graph(wb + p, t, { method: "PATCH", headers: h, body: JSON.stringify(body) }) });
  } finally {
    graph(wb + "/closeSession", t, { method: "POST", headers: h }).catch(() => {
    });
  }
}
async function tableLayout(s2) {
  if (layout) return layout;
  const r = await s2.get(`/tables/${TABLE()}/range?$select=address`);
  const m = /^(.*)!\$?([A-Z]+)\$?(\d+):/.exec(r.address);
  layout = { sheet: m[1].replace(/^'|'$/g, ""), col: m[2], headerRow: parseInt(m[3], 10) };
  return layout;
}
async function readTable(s2) {
  const [hdr, body] = await Promise.all([s2.get(`/tables/${TABLE()}/headerRowRange?$select=values`), s2.get(`/tables/${TABLE()}/rows?$select=values`)]);
  const headers = hdr.values[0].map((h) => String(h).trim());
  return { headers, rows: body.value.map((x, i) => ({ i, raw: x.values[0], obj: Object.fromEntries(headers.map((h, k) => [h, x.values[0][k]])) })) };
}
var colLetter = (n2) => {
  let s2 = "";
  n2 += 1;
  while (n2 > 0) {
    const m = (n2 - 1) % 26;
    s2 = String.fromCharCode(65 + m) + s2;
    n2 = Math.floor((n2 - 1) / 26);
  }
  return s2;
};
var serial = (iso) => iso ? Math.round(Date.parse(String(iso).slice(0, 10) + "T00:00:00Z") / 864e5 + 25569) : "";
var fromSerial = (v) => typeof v === "number" && v > 2e4 ? new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10) : typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v) ? v.slice(0, 10) : null;

// lib/wati.js
async function send(mobile, template, parameters) {
  const base = process.env.WATI_API_URL, key = process.env.WATI_API_KEY;
  if (!base || !key) return { skipped: "WATI not configured" };
  const num = String(mobile).replace(/\D/g, "").replace(/^0+/, "");
  const wa = num.length === 10 ? "91" + num : num;
  const r = await fetch(`${base.replace(/\/$/, "")}/api/v1/sendTemplateMessage?whatsappNumber=${wa}`, {
    method: "POST",
    headers: { Authorization: "Bearer " + key, "Content-Type": "application/json" },
    body: JSON.stringify({ template_name: template, broadcast_name: template, parameters })
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || j.result === false) throw new Error("WATI: " + (j.info || j.message || r.status));
  return { ok: true };
}
var sendDispatchLink = ({ mobile, name, coils, doNo, link }) => send(mobile, process.env.WATI_TEMPLATE || "nmdc_dispatch_link", [{ name: "name", value: name || "Partner" }, { name: "coils", value: String(coils) }, { name: "do_no", value: doNo || "-" }, { name: "link", value: link }]);
var sendTransporterLink = ({ mobile, name, link }) => send(mobile, process.env.WATI_TEMPLATE_TRANSPORTER || "nmdc_transporter_link", [{ name: "name", value: name || "Partner" }, { name: "link", value: link }]);

// api/sync.js
var OWN = ["Entry Date", "NMDC SO No", "NMDC Invoice / DO No", "Coil No", "Form", "Grade", "Thk (mm)", "Width (mm)", "Length (mm)", "Weight (MT)", "Destination", "Transporter Name", "Transporter Mobile", "Remarks"];
var MIRROR = ["Status", "Vehicle No", "Transporter Name", "Transporter Mobile", "LR No", "Dispatch Date (ex NMDC)", "Expected Arrival", "Actual Arrival"];
var s = (v) => v == null ? "" : String(v).trim();
var n = (v) => typeof v === "number" ? v : v === "" || v == null ? null : Number(v);
var hash = (o) => crypto.createHash("sha1").update(OWN.map((k) => s(o[k])).join("|")).digest("hex");
async function handler2(req, res) {
  const auth = req.headers.authorization || "";
  if (!process.env.CRON_SECRET) return json(res, 503, { error: "CRON_SECRET not set" });
  const key = decodeURIComponent((/[?&]key=([^&]*)/.exec(req.url || "") || [])[1] || req.query && req.query.key || "");
  if (auth !== "Bearer " + process.env.CRON_SECRET && key !== process.env.CRON_SECRET)
    return json(res, 401, { error: "unauthorized", hint: { url_seen: req.url, key_received_length: key.length, secret_configured_length: (process.env.CRON_SECRET || "").length, header_present: !!auth } });
  const out = { transporters: 0, imported: 0, updated: 0, dispatches: 0, notified: 0, mirrored: 0, appended: 0, errors: [] };
  const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
  try {
    await withSession(true, async (sess) => {
      try {
        const tr = await sess.get(`/tables/tbl_Transporters/rows?$select=values`);
        for (const [i, row] of tr.value.entries()) {
          const [name, mobile, link, sent, active] = row.values[0].map((v) => s(v));
          const mob = mobile.replace(/\D/g, "");
          if (!name || mob.length < 10) continue;
          let t = (await sql`select * from transporters where mobile=${mob}`)[0];
          if (!t) {
            const token2 = crypto.randomBytes(9).toString("base64url");
            t = (await sql`insert into transporters (name, mobile, token, active, excel_row) values (${name}, ${mob}, ${token2}, ${active.toLowerCase() !== "no"}, ${i}) returning *`)[0];
            out.transporters++;
          } else if (t.name !== name || t.active !== (active.toLowerCase() !== "no") || t.excel_row !== i) {
            await sql`update transporters set name=${name}, active=${active.toLowerCase() !== "no"}, excel_row=${i} where id=${t.id}`;
          }
          const wantLink = `${base}/t/${t.token}`;
          let sentVal = sent;
          if (!t.link_sent_at && active.toLowerCase() !== "no") {
            try {
              const r = await sendTransporterLink({ mobile: mob, name, link: wantLink });
              if (!r.skipped) {
                await sql`update transporters set link_sent_at=now(), notify_error=null where id=${t.id}`;
                sentVal = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
                out.notified++;
              } else await sql`update transporters set notify_error=${r.skipped} where id=${t.id}`;
            } catch (e) {
              out.errors.push(e.message);
              await sql`update transporters set notify_error=${e.message} where id=${t.id}`;
            }
          }
          if (link !== wantLink || sentVal !== sent) await sess.patch(`/tables/tbl_Transporters/rows/itemAt(index=${i})`, { values: [[name, mobile, wantLink, sentVal, active || "Yes"]] });
        }
      } catch (e) {
        if (!/404|ItemNotFound/i.test(e.message)) throw e;
      }
      const tmap = Object.fromEntries((await sql`select id, name, mobile from transporters`).map((t) => [t.mobile, t]));
      const tByName = Object.fromEntries(Object.values(tmap).map((t) => [t.name.toLowerCase(), t]));
      const { headers, rows } = await readTable(sess);
      const hidx = Object.fromEntries(headers.map((h, i) => [h, i]));
      for (const k of [...OWN, ...MIRROR, "Entered By"]) if (!(k in hidx)) throw new Error(`Column "${k}" not found in ${TABLE()}`);
      const live = rows.filter((r) => s(r.obj["Coil No"]) !== "");
      const existing = Object.fromEntries((await sql`select coil_no, excel_hash, excel_row, dispatch_id, status, source, updated_at, mirrored_at from coils`).map((c) => [c.coil_no, c]));
      const fresh = [];
      for (const r of live) {
        const o = r.obj, coil = s(o["Coil No"]), h = hash(o), ex = existing[coil];
        if (ex && ex.source === "transporter") {
          if (ex.excel_row !== r.i) await sql`update coils set excel_row=${r.i} where coil_no=${coil}`;
          continue;
        }
        const own = {
          entry_date: fromSerial(o["Entry Date"]),
          so_no: s(o["NMDC SO No"]),
          do_no: s(o["NMDC Invoice / DO No"]),
          form: s(o["Form"]),
          grade: s(o["Grade"]),
          thk: n(o["Thk (mm)"]),
          width: n(o["Width (mm)"]),
          length: n(o["Length (mm)"]),
          weight: n(o["Weight (MT)"]),
          destination: s(o["Destination"]),
          transporter_name: s(o["Transporter Name"]),
          transporter_mobile: s(o["Transporter Mobile"]).replace(/\D/g, ""),
          remarks: s(o["Remarks"])
        };
        if (!own.transporter_mobile && own.transporter_name && tByName[own.transporter_name.toLowerCase()]) own.transporter_mobile = tByName[own.transporter_name.toLowerCase()].mobile;
        const tid = tmap[own.transporter_mobile]?.id ?? null;
        if (!ex) {
          const st = STATUSES.includes(s(o["Status"])) ? s(o["Status"]) : "Under Loading";
          await sql`insert into coils (coil_no, entry_date, so_no, do_no, form, grade, thk, width, length, weight, destination, transporter_name, transporter_mobile, remarks, status,
              vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, excel_row, excel_hash, mirrored_at, source, transporter_id, entered_by)
            values (${coil}, ${own.entry_date}, ${own.so_no}, ${own.do_no}, ${own.form}, ${own.grade}, ${own.thk}, ${own.width}, ${own.length}, ${own.weight}, ${own.destination},
              ${own.transporter_name}, ${own.transporter_mobile}, ${own.remarks}, ${st}, ${s(o["Vehicle No"])}, ${s(o["LR No"])}, ${fromSerial(o["Dispatch Date (ex NMDC)"])},
              ${fromSerial(o["Expected Arrival"])}, ${fromSerial(o["Actual Arrival"])}, ${r.i}, ${h}, now(), 'bsc', ${tid}, ${s(o["Entered By"]) || "BSC"})`;
          await sql`insert into events (coil_no, action, actor, detail) values (${coil}, 'imported', 'sync', ${JSON.stringify({ row: r.i })})`;
          out.imported++;
          fresh.push({ coil, ...own });
        } else if (ex.excel_hash !== h || ex.excel_row !== r.i) {
          await sql`update coils set entry_date=${own.entry_date}, so_no=${own.so_no}, do_no=${own.do_no}, form=${own.form}, grade=${own.grade}, thk=${own.thk}, width=${own.width}, length=${own.length},
              weight=${own.weight}, destination=${own.destination}, transporter_name=${own.transporter_name}, transporter_mobile=${own.transporter_mobile}, remarks=${own.remarks},
              excel_row=${r.i}, excel_hash=${h}, transporter_id=${tid} where coil_no=${coil}`;
          if (ex.excel_hash !== h) out.updated++;
          if (!ex.dispatch_id && own.transporter_mobile) fresh.push({ coil, ...own });
        }
      }
      const groups = {};
      for (const f of fresh) {
        if (!f.transporter_mobile) continue;
        const k = f.do_no + "|" + f.transporter_mobile;
        (groups[k] ||= []).push(f);
      }
      for (const k of Object.keys(groups)) {
        const g = groups[k], f = g[0];
        let disp = (await sql`select * from dispatches where do_no=${f.do_no} and transporter_mobile=${f.transporter_mobile} order by id desc limit 1`)[0];
        if (!disp) {
          const token2 = crypto.randomBytes(9).toString("base64url");
          disp = (await sql`insert into dispatches (token, do_no, transporter_name, transporter_mobile, transporter_id) values (${token2}, ${f.do_no}, ${f.transporter_name}, ${f.transporter_mobile}, ${tmap[f.transporter_mobile]?.id ?? null}) returning *`)[0];
          out.dispatches++;
        }
        await sql`update coils set dispatch_id=${disp.id} where coil_no = any(${g.map((x) => x.coil)})`;
        try {
          const count = (await sql`select count(*)::int as c from coils where dispatch_id=${disp.id}`)[0].c;
          const standing = tmap[disp.transporter_mobile];
          const r = await sendDispatchLink({ mobile: disp.transporter_mobile, name: disp.transporter_name, coils: count, doNo: disp.do_no, link: `${base}/t/${standing ? standing.token : disp.token}` });
          await sql`update dispatches set notified_at=now(), notify_error=${r.skipped || null} where id=${disp.id}`;
          if (!r.skipped) out.notified++;
        } catch (e) {
          out.errors.push(e.message);
          await sql`update dispatches set notify_error=${e.message} where id=${disp.id}`;
        }
      }
      const fullRow = (c) => headers.map((h) => ({
        "Sl No": "",
        "Entry Date": serial(c.entry_date),
        "NMDC SO No": c.so_no || "",
        "NMDC Invoice / DO No": c.do_no || "",
        "Coil No": c.coil_no,
        "Form": c.form || "",
        "Grade": c.grade || "",
        "Thk (mm)": c.thk ?? "",
        "Width (mm)": c.width ?? "",
        "Length (mm)": c.length ?? "",
        "Weight (MT)": c.weight ?? "",
        "Status": c.status,
        "Vehicle No": c.vehicle_no || "",
        "Transporter Name": c.transporter_name || "",
        "Transporter Mobile": c.transporter_mobile || "",
        "LR No": c.lr_no || "",
        "Dispatch Date (ex NMDC)": serial(c.dispatch_date),
        "Expected Arrival": serial(c.expected_arrival),
        "Actual Arrival": serial(c.actual_arrival),
        "Destination": c.destination || "",
        "Remarks": c.remarks || "",
        "Entered By": c.entered_by || c.transporter_name || "",
        "Last Updated": ""
      })[h] ?? "");
      const news = await sql`select * from coils where source='transporter' and excel_row is null order by created_at`;
      if (news.length) {
        const lay = await tableLayout(sess);
        const firstEmpty = live.length;
        for (const [k, c] of news.entries()) {
          const rowIdx = firstEmpty + k, sheetRow = lay.headerRow + 1 + rowIdx;
          if (rowIdx >= rows.length) {
            out.errors.push("tbl_Dispatch is full \u2014 extend the table");
            break;
          }
          const vals = fullRow(c);
          vals[headers.indexOf("Sl No")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",ROW()-${lay.headerRow})`;
          vals[headers.indexOf("Last Updated")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",NOW())`;
          const fmt = headers.map((h) => /Date|Arrival|Updated/.test(h) ? "dd-mmm-yy" : /Weight/.test(h) ? "#,##0.000" : /Mobile|No$|Coil No|LR No/.test(h) ? "@" : "General");
          await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${colLetter(0)}${sheetRow}:${colLetter(headers.length - 1)}${sheetRow}')`, { values: [vals], numberFormat: [fmt] });
          await sql`update coils set excel_row=${rowIdx}, mirrored_at=now() where coil_no=${c.coil_no}`;
          live.push({ i: rowIdx, obj: { "Coil No": c.coil_no } });
          out.appended++;
        }
      }
      const dirty = await sql`select * from coils where excel_row is not null and (mirrored_at is null or updated_at > mirrored_at)`;
      if (dirty.length) {
        const lay = await tableLayout(sess);
        const firstCol = headers.indexOf("Status"), lastCol = headers.indexOf("Actual Arrival");
        const byRow = Object.fromEntries(live.map((r) => [r.i, r]));
        for (const c of dirty) {
          const r = byRow[c.excel_row];
          if (!r || s(r.obj["Coil No"]) !== c.coil_no) {
            out.errors.push(`row moved for ${c.coil_no}`);
            continue;
          }
          const sheetRow = lay.headerRow + 1 + c.excel_row;
          if (c.source === "transporter") {
            const vals2 = fullRow(c);
            vals2[headers.indexOf("Sl No")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",ROW()-${lay.headerRow})`;
            vals2[headers.indexOf("Last Updated")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",NOW())`;
            await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${colLetter(0)}${sheetRow}:${colLetter(headers.length - 1)}${sheetRow}')`, { values: [vals2] });
            await sql`update coils set mirrored_at=now() where coil_no=${c.coil_no}`;
            out.mirrored++;
            continue;
          }
          const addr = `${colLetter(firstCol)}${sheetRow}:${colLetter(lastCol)}${sheetRow}`;
          const vals = headers.slice(firstCol, lastCol + 1).map((h) => ({
            "Status": c.status,
            "Vehicle No": c.vehicle_no || "",
            "Transporter Name": c.transporter_name || "",
            "Transporter Mobile": c.transporter_mobile || "",
            "LR No": c.lr_no || "",
            "Dispatch Date (ex NMDC)": serial(c.dispatch_date),
            "Expected Arrival": serial(c.expected_arrival),
            "Actual Arrival": serial(c.actual_arrival)
          })[h]);
          const fmt = headers.slice(firstCol, lastCol + 1).map((h) => /Date|Arrival/.test(h) ? "dd-mmm-yy" : "@");
          await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${addr}')`, { values: [vals], numberFormat: [fmt] });
          await sql`update coils set mirrored_at=now() where coil_no=${c.coil_no}`;
          out.mirrored++;
        }
      }
    });
    json(res, 200, { ok: true, ...out });
  } catch (e) {
    json(res, 500, { ok: false, ...out, error: e.message });
  }
}

// api/tform.js
import crypto2 from "node:crypto";
var VEH = /^[A-Z]{2}[ -]?\d{1,2}[ -]?[A-Z]{0,3}[ -]?\d{4}$/i;
var destCache = { at: 0, list: [] };
async function destinations() {
  if (Date.now() - destCache.at < 10 * 6e4 && destCache.list.length) return destCache.list;
  try {
    const list = await withSession(false, async (s2) => (await s2.get(`/worksheets('Lists')/range(address='D2:D31')?$select=values`)).values.flat().map((v) => String(v).trim()).filter(Boolean));
    destCache = { at: Date.now(), list };
  } catch {
  }
  return destCache.list;
}
var group = (coils) => {
  const g = {};
  for (const c of coils) (g[c.do_no || "(no DO)"] ||= []).push(c);
  return Object.entries(g).map(([do_no, coils2]) => ({ do_no, coils: coils2 }));
};
async function handler3(req, res) {
  const qtoken = decodeURIComponent((/[?&]token=([^&]*)/.exec(req.url || "") || [])[1] || req.query && req.query.token || "");
  const token2 = String((req.method === "GET" ? qtoken : req.body?.token) || "");
  if (!token2) return json(res, 400, { error: "token missing" });
  let tp = (await sql`select * from transporters where token=${token2} and active`)[0];
  let disp = null;
  if (!tp) {
    disp = (await sql`select * from dispatches where token=${token2}`)[0];
    if (!disp) return json(res, 404, { error: "This link is not valid. Please contact Bharat Steel." });
    if (disp.transporter_id) tp = (await sql`select * from transporters where id=${disp.transporter_id} and active`)[0] || null;
  }
  const name = tp?.name || disp?.transporter_name || "";
  const mobile = tp?.mobile || disp?.transporter_mobile || "";
  const actor = "transporter:" + mobile;
  const mine = () => tp ? sql`select coil_no, do_no, so_no, form, weight, destination, status, vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, source, entered_by
          from coils where transporter_id=${tp.id} or transporter_mobile=${tp.mobile} order by created_at desc, coil_no` : sql`select coil_no, do_no, so_no, form, weight, destination, status, vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, source, entered_by
          from coils where dispatch_id=${disp.id} order by coil_no`;
  if (req.method === "GET") return json(res, 200, { mode: tp ? "transporter" : "dispatch", name, do_no: disp?.do_no || null, groups: group(await mine()), destinations: await destinations() });
  if (req.method !== "POST") return json(res, 405, { error: "method" });
  const b = req.body || {}, action = String(b.action || ""), today = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
  if (action === "create") {
    if (!tp) return json(res, 403, { error: "This link cannot create new dispatches" });
    const doNo = String(b.do_no || "").trim(), soNo = String(b.so_no || "").trim(), dest = String(b.destination || "").trim();
    const veh = String(b.vehicle_no || "").toUpperCase().replace(/\s+/g, " ").trim(), lr = String(b.lr_no || "").trim();
    const dd = d(b.dispatch_date) || today, ea = d(b.expected_arrival);
    const coils = (Array.isArray(b.coils) ? b.coils : []).map((c) => ({ coil_no: String(c.coil_no || "").trim().toUpperCase(), weight: Number(c.weight), form: c.form === "CTL Sheet" ? "CTL Sheet" : "Coil" })).filter((c) => c.coil_no);
    if (!doNo) return json(res, 400, { error: "Enter the NMDC DO / invoice number" });
    if (!dest) return json(res, 400, { error: "Select the destination" });
    if (!coils.length) return json(res, 400, { error: "Add at least one coil" });
    if (coils.some((c) => !(c.weight > 0))) return json(res, 400, { error: "Enter the weight (MT) for every coil" });
    if (veh && !VEH.test(veh)) return json(res, 400, { error: "Enter a valid vehicle number, e.g. TN 20 BX 7719" });
    if (veh && !ea) return json(res, 400, { error: "Enter the expected arrival date" });
    const dup = await sql`select coil_no from coils where coil_no = any(${coils.map((c) => c.coil_no)})`;
    if (dup.length) return json(res, 400, { error: "Already logged: " + dup.map((x) => x.coil_no).join(", ") });
    const status = veh ? "Loaded" : "Under Loading";
    let dsp = (await sql`select * from dispatches where do_no=${doNo} and transporter_mobile=${tp.mobile} order by id desc limit 1`)[0];
    if (!dsp) dsp = (await sql`insert into dispatches (token, do_no, transporter_name, transporter_mobile, transporter_id, notified_at) values (${crypto2.randomBytes(9).toString("base64url")}, ${doNo}, ${tp.name}, ${tp.mobile}, ${tp.id}, now()) returning *`)[0];
    for (const c of coils) {
      await sql`insert into coils (coil_no, entry_date, so_no, do_no, form, weight, destination, transporter_name, transporter_mobile, transporter_id, status, vehicle_no, lr_no, dispatch_date, expected_arrival, dispatch_id, source, entered_by)
                values (${c.coil_no}, ${today}, ${soNo}, ${doNo}, ${c.form}, ${c.weight}, ${dest}, ${tp.name}, ${tp.mobile}, ${tp.id}, ${status}, ${veh}, ${lr}, ${veh ? dd : null}, ${ea}, ${dsp.id}, 'transporter', ${tp.name})`;
      await sql`insert into events (coil_no, action, actor, detail) values (${c.coil_no}, 'created', ${actor}, ${JSON.stringify({ doNo, dest, veh, lr })})`;
    }
    return json(res, 200, { ok: true, groups: group(await mine()) });
  }
  const sel = Array.isArray(b.coils) ? b.coils.map(String) : [];
  if (!sel.length) return json(res, 400, { error: "Select at least one coil" });
  const own = tp ? await sql`select coil_no, status, source from coils where coil_no = any(${sel}) and (transporter_id=${tp.id} or transporter_mobile=${tp.mobile})` : await sql`select coil_no, status, source from coils where coil_no = any(${sel}) and dispatch_id=${disp.id}`;
  if (own.length !== sel.length) return json(res, 403, { error: "One of the coils is not on this link" });
  if (action === "load") {
    const veh = String(b.vehicle_no || "").toUpperCase().replace(/\s+/g, " ").trim(), lr = String(b.lr_no || "").trim();
    const dd = d(b.dispatch_date) || today, ea = d(b.expected_arrival);
    if (!VEH.test(veh)) return json(res, 400, { error: "Enter a valid vehicle number, e.g. TN 20 BX 7719" });
    if (!ea) return json(res, 400, { error: "Enter the expected arrival date" });
    await sql`update coils set vehicle_no=${veh}, lr_no=${lr}, dispatch_date=${dd}, expected_arrival=${ea}, status = case when status='Under Loading' then 'Loaded' else status end, updated_at=now() where coil_no = any(${sel}) and status <> 'Arrived'`;
    await sql`insert into events (coil_no, action, actor, detail) select unnest(${sel}::text[]), 'load', ${actor}, ${JSON.stringify({ veh, lr, dd, ea })}`;
  } else if (action === "gateout") {
    const blocked = own.filter((c) => c.status === "Under Loading");
    if (blocked.length) return json(res, 400, { error: "Enter truck details first for " + blocked.map((c) => c.coil_no).join(", ") });
    await sql`update coils set status='In-Transit', dispatch_date=coalesce(dispatch_date, ${today}), updated_at=now() where coil_no = any(${sel}) and status='Loaded'`;
    await sql`insert into events (coil_no, action, actor) select unnest(${sel}::text[]), 'gateout', ${actor}`;
  } else if (action === "delivered") {
    await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${d(b.actual_arrival) || today}), updated_at=now() where coil_no = any(${sel}) and status in ('Loaded','In-Transit')`;
    await sql`insert into events (coil_no, action, actor) select unnest(${sel}::text[]), 'delivered', ${actor}`;
  } else if (action === "edit") {
    const editable = own.filter((c) => c.source === "transporter" && c.status !== "Arrived");
    if (editable.length !== own.length) return json(res, 403, { error: "Only your own undelivered entries can be edited" });
    const w = b.weight == null ? null : Number(b.weight), dest = b.destination == null ? null : String(b.destination).trim();
    await sql`update coils set weight=coalesce(${w}, weight), destination=coalesce(${dest}, destination), updated_at=now() where coil_no = any(${sel})`;
    await sql`insert into events (coil_no, action, actor, detail) select unnest(${sel}::text[]), 'edit', ${actor}, ${JSON.stringify({ w, dest })}`;
  } else return json(res, 400, { error: "unknown action" });
  json(res, 200, { ok: true, groups: group(await mine()) });
}

// api/arrive.js
async function handler4(req, res) {
  if (cors(req, res)) return;
  if (req.method !== "POST") return json(res, 405, { error: "method" });
  const b = req.body || {};
  if (!process.env.YARD_PIN) return json(res, 503, { error: "YARD_PIN not set" });
  if (String(b.pin || "") !== process.env.YARD_PIN) return json(res, 401, { error: "Wrong PIN" });
  const today = (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
  let rows;
  if (b.vehicle_no) rows = await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${today}), updated_at=now()
                                       where upper(regexp_replace(vehicle_no,'\s','','g')) = upper(regexp_replace(${String(b.vehicle_no)},'\s','','g')) and status in ('Loaded','In-Transit') returning coil_no`;
  else if (Array.isArray(b.coils)) rows = await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${today}), updated_at=now() where coil_no = any(${b.coils.map(String)}) and status in ('Loaded','In-Transit') returning coil_no`;
  else return json(res, 400, { error: "vehicle_no or coils required" });
  for (const r of rows) await sql`insert into events (coil_no, action, actor) values (${r.coil_no}, 'delivered', 'yard')`;
  json(res, 200, { ok: true, updated: rows.map((r) => r.coil_no) });
}

// _router.js
async function handler5(req, res) {
  const path = (req.url || "").split("?")[0].replace(/\/+$/, "");
  const route = path.replace(/^\/api\/?/, "");
  if (route === "dispatch") return handler(req, res);
  if (route === "sync") return handler2(req, res);
  if (route === "tform") return handler3(req, res);
  if (route === "arrive") return handler4(req, res);
  res.setHeader("Cache-Control", "no-store");
  res.status(404).json({ error: "unknown route: " + route });
}
export {
  handler5 as default
};
