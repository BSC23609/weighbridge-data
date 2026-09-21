// Transporter API. Token is either a per-DO dispatch link (mode 1) or a transporter's standing link (mode 1 + 2).
//   GET  ?token=...                                -> { mode, name, do_no?, groups:[{do_no, coils:[...]}], destinations }
//   POST { token, action, ... }                    -> actions: load | gateout | delivered | create | edit
import crypto from "node:crypto";
import { sql, json, d } from "../lib/db.js";
import { withSession } from "../lib/graph.js";
const VEH = /^[A-Z]{2}[ -]?\d{1,2}[ -]?[A-Z]{0,3}[ -]?\d{4}$/i;
let destCache = { at: 0, list: [] };

async function destinations() {
  if (Date.now() - destCache.at < 10 * 60000 && destCache.list.length) return destCache.list;
  try {
    const list = await withSession(false, async (s) => (await s.get(`/worksheets('Lists')/range(address='D2:D31')?$select=values`)).values.flat().map((v) => String(v).trim()).filter(Boolean));
    destCache = { at: Date.now(), list };
  } catch { /* keep old */ }
  return destCache.list;
}
const group = (coils) => { const g = {}; for (const c of coils) (g[c.do_no || "(no DO)"] ||= []).push(c); return Object.entries(g).map(([do_no, coils]) => ({ do_no, coils })); };

export default async function handler(req, res) {
  const token = String((req.method === "GET" ? req.query.token : req.body?.token) || "");
  if (!token) return json(res, 400, { error: "token missing" });

  // resolve who this is
  let tp = (await sql`select * from transporters where token=${token} and active`)[0];
  let disp = null;
  if (!tp) {
    disp = (await sql`select * from dispatches where token=${token}`)[0];
    if (!disp) return json(res, 404, { error: "This link is not valid. Please contact Bharat Steel." });
    if (disp.transporter_id) tp = (await sql`select * from transporters where id=${disp.transporter_id} and active`)[0] || null;
  }
  const name = tp?.name || disp?.transporter_name || "";
  const mobile = tp?.mobile || disp?.transporter_mobile || "";
  const actor = "transporter:" + mobile;
  // coils visible on this link
  const mine = () => tp
    ? sql`select coil_no, do_no, so_no, form, weight, destination, status, vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, source, entered_by
          from coils where transporter_id=${tp.id} or transporter_mobile=${tp.mobile} order by created_at desc, coil_no`
    : sql`select coil_no, do_no, so_no, form, weight, destination, status, vehicle_no, lr_no, dispatch_date, expected_arrival, actual_arrival, source, entered_by
          from coils where dispatch_id=${disp.id} order by coil_no`;

  if (req.method === "GET") return json(res, 200, { mode: tp ? "transporter" : "dispatch", name, do_no: disp?.do_no || null, groups: group(await mine()), destinations: await destinations() });
  if (req.method !== "POST") return json(res, 405, { error: "method" });

  const b = req.body || {}, action = String(b.action || ""), today = new Date().toISOString().slice(0, 10);

  if (action === "create") {                                   // mode 2 — transporter enters everything
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
    if (!dsp) dsp = (await sql`insert into dispatches (token, do_no, transporter_name, transporter_mobile, transporter_id, notified_at) values (${crypto.randomBytes(9).toString("base64url")}, ${doNo}, ${tp.name}, ${tp.mobile}, ${tp.id}, now()) returning *`)[0];
    for (const c of coils) {
      await sql`insert into coils (coil_no, entry_date, so_no, do_no, form, weight, destination, transporter_name, transporter_mobile, transporter_id, status, vehicle_no, lr_no, dispatch_date, expected_arrival, dispatch_id, source, entered_by)
                values (${c.coil_no}, ${today}, ${soNo}, ${doNo}, ${c.form}, ${c.weight}, ${dest}, ${tp.name}, ${tp.mobile}, ${tp.id}, ${status}, ${veh}, ${lr}, ${veh ? dd : null}, ${ea}, ${dsp.id}, 'transporter', ${tp.name})`;
      await sql`insert into events (coil_no, action, actor, detail) values (${c.coil_no}, 'created', ${actor}, ${JSON.stringify({ doNo, dest, veh, lr })})`;
    }
    return json(res, 200, { ok: true, groups: group(await mine()) });
  }

  const sel = Array.isArray(b.coils) ? b.coils.map(String) : [];
  if (!sel.length) return json(res, 400, { error: "Select at least one coil" });
  const own = tp
    ? await sql`select coil_no, status, source from coils where coil_no = any(${sel}) and (transporter_id=${tp.id} or transporter_mobile=${tp.mobile})`
    : await sql`select coil_no, status, source from coils where coil_no = any(${sel}) and dispatch_id=${disp.id}`;
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
  } else if (action === "edit") {                               // mode 2 only: correct own entry (weight / destination / DO) before delivery
    const editable = own.filter((c) => c.source === "transporter" && c.status !== "Arrived");
    if (editable.length !== own.length) return json(res, 403, { error: "Only your own undelivered entries can be edited" });
    const w = b.weight == null ? null : Number(b.weight), dest = b.destination == null ? null : String(b.destination).trim();
    await sql`update coils set weight=coalesce(${w}, weight), destination=coalesce(${dest}, destination), updated_at=now() where coil_no = any(${sel})`;
    await sql`insert into events (coil_no, action, actor, detail) select unnest(${sel}::text[]), 'edit', ${actor}, ${JSON.stringify({ w, dest })}`;
  } else return json(res, 400, { error: "unknown action" });
  json(res, 200, { ok: true, groups: group(await mine()) });
}
