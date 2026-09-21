// Dashboard data from Neon, returned with the same header names the page already understands.
import { sql, json, cors } from "../lib/db.js";
export default async function handler(req, res) {
  if (cors(req, res)) return;
  if (!process.env.DASH_PIN) return json(res, 503, { error: "DASH_PIN not set" });
  if (req.headers["x-dash-pin"] !== process.env.DASH_PIN) return json(res, 401, { error: "pin" });
  try {
    const c = await sql`select * from coils order by entry_date nulls first, created_at`;
    const rows = c.map((r) => ({ "Entry Date": r.entry_date, "NMDC SO No": r.so_no, "NMDC Invoice / DO No": r.do_no, "Coil No": r.coil_no, "Form": r.form, "Grade": r.grade,
      "Thk (mm)": r.thk == null ? "" : Number(r.thk), "Width (mm)": r.width ?? "", "Length (mm)": r.length ?? "", "Weight (MT)": r.weight == null ? "" : Number(r.weight),
      "Status": r.status, "Vehicle No": r.vehicle_no || "", "Transporter Name": r.transporter_name || "", "Transporter Mobile": r.transporter_mobile || "", "LR No": r.lr_no || "",
      "Dispatch Date (ex NMDC)": r.dispatch_date, "Expected Arrival": r.expected_arrival, "Actual Arrival": r.actual_arrival, "Destination": r.destination || "", "Remarks": r.remarks || "", "Entered By": r.entered_by || "BSC" }));
    json(res, 200, { rows, fetchedAt: new Date().toISOString() });
  } catch (e) { json(res, 502, { error: e.message }); }
}
