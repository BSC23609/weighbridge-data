// Yard marks a vehicle's coils as Arrived from the dashboard. POST {pin, vehicle_no} or {pin, coils:[...]}
import { sql, json, cors } from "../lib/db.js";
export default async function handler(req, res) {
  if (cors(req, res)) return;
  if (req.method !== "POST") return json(res, 405, { error: "method" });
  const b = req.body || {};
  if (!process.env.YARD_PIN) return json(res, 503, { error: "YARD_PIN not set" });
  if (String(b.pin || "") !== process.env.YARD_PIN) return json(res, 401, { error: "Wrong PIN" });
  const today = new Date().toISOString().slice(0, 10);
  let rows;
  if (b.vehicle_no) rows = await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${today}), updated_at=now()
                                       where upper(regexp_replace(vehicle_no,'\s','','g')) = upper(regexp_replace(${String(b.vehicle_no)},'\s','','g')) and status in ('Loaded','In-Transit') returning coil_no`;
  else if (Array.isArray(b.coils)) rows = await sql`update coils set status='Arrived', actual_arrival=coalesce(actual_arrival, ${today}), updated_at=now() where coil_no = any(${b.coils.map(String)}) and status in ('Loaded','In-Transit') returning coil_no`;
  else return json(res, 400, { error: "vehicle_no or coils required" });
  for (const r of rows) await sql`insert into events (coil_no, action, actor) values (${r.coil_no}, 'delivered', 'yard')`;
  json(res, 200, { ok: true, updated: rows.map((r) => r.coil_no) });
}
