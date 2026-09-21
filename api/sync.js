// Excel <-> Neon sync. Call every 5 minutes from cron-job.org with header  Authorization: Bearer <CRON_SECRET>
//   1. Import new / changed coils from tbl_Dispatch (BSC-owned fields) into Neon.
//   2. Group new coils by DO + transporter mobile into a dispatch, create its link, send WhatsApp via WATI.
//   3. Mirror transporter-owned fields (status, vehicle, LR, dates) from Neon back into the same Excel rows.
import crypto from "node:crypto";
import { sql, json, STATUSES } from "../lib/db.js";
import { withSession, readTable, tableLayout, TABLE, colLetter, serial, fromSerial } from "../lib/graph.js";
import { sendDispatchLink, sendTransporterLink } from "../lib/wati.js";

const OWN = ["Entry Date", "NMDC SO No", "NMDC Invoice / DO No", "Coil No", "Form", "Grade", "Thk (mm)", "Width (mm)", "Length (mm)", "Weight (MT)", "Destination", "Transporter Name", "Transporter Mobile", "Remarks"];
const MIRROR = ["Status", "Vehicle No", "Transporter Name", "Transporter Mobile", "LR No", "Dispatch Date (ex NMDC)", "Expected Arrival", "Actual Arrival"];
const s = (v) => (v == null ? "" : String(v).trim());
const n = (v) => (typeof v === "number" ? v : (v === "" || v == null ? null : Number(v)));
const hash = (o) => crypto.createHash("sha1").update(OWN.map((k) => s(o[k])).join("|")).digest("hex");

export default async function handler(req, res) {
  const auth = req.headers.authorization || "";
  if (!process.env.CRON_SECRET) return json(res, 503, { error: "CRON_SECRET not set" });
  if (auth !== "Bearer " + process.env.CRON_SECRET) return json(res, 401, { error: "unauthorized" });
  const out = { transporters: 0, imported: 0, updated: 0, dispatches: 0, notified: 0, mirrored: 0, appended: 0, errors: [] };
  const base = (process.env.PUBLIC_URL || `https://${req.headers.host}`).replace(/\/$/, "");
  try {
    await withSession(true, async (sess) => {
      // 0. transporter registry (tbl_Transporters: Name, Mobile, Link, Link Sent, Active)
      try {
        const tr = await sess.get(`/tables/tbl_Transporters/rows?$select=values`);
        for (const [i, row] of tr.value.entries()) {
          const [name, mobile, link, sent, active] = row.values[0].map((v) => s(v));
          const mob = mobile.replace(/\D/g, "");
          if (!name || mob.length < 10) continue;
          let t = (await sql`select * from transporters where mobile=${mob}`)[0];
          if (!t) {
            const token = crypto.randomBytes(9).toString("base64url");
            t = (await sql`insert into transporters (name, mobile, token, active, excel_row) values (${name}, ${mob}, ${token}, ${active.toLowerCase() !== "no"}, ${i}) returning *`)[0];
            out.transporters++;
          } else if (t.name !== name || t.active !== (active.toLowerCase() !== "no") || t.excel_row !== i) {
            await sql`update transporters set name=${name}, active=${active.toLowerCase() !== "no"}, excel_row=${i} where id=${t.id}`;
          }
          const wantLink = `${base}/t/${t.token}`;
          let sentVal = sent;
          if (!t.link_sent_at && (active.toLowerCase() !== "no")) {
            try { const r = await sendTransporterLink({ mobile: mob, name, link: wantLink }); if (!r.skipped) { await sql`update transporters set link_sent_at=now(), notify_error=null where id=${t.id}`; sentVal = new Date().toISOString().slice(0, 10); out.notified++; } else await sql`update transporters set notify_error=${r.skipped} where id=${t.id}`; }
            catch (e) { out.errors.push(e.message); await sql`update transporters set notify_error=${e.message} where id=${t.id}`; }
          }
          if (link !== wantLink || sentVal !== sent) await sess.patch(`/tables/tbl_Transporters/rows/itemAt(index=${i})`, { values: [[name, mobile, wantLink, sentVal, active || "Yes"]] });
        }
      } catch (e) { if (!/404|ItemNotFound/i.test(e.message)) throw e; }   // sheet not present yet -> skip
      const tmap = Object.fromEntries((await sql`select id, name, mobile from transporters`).map((t) => [t.mobile, t]));
      const tByName = Object.fromEntries(Object.values(tmap).map((t) => [t.name.toLowerCase(), t]));

      const { headers, rows } = await readTable(sess);
      const hidx = Object.fromEntries(headers.map((h, i) => [h, i]));
      for (const k of [...OWN, ...MIRROR, "Entered By"]) if (!(k in hidx)) throw new Error(`Column "${k}" not found in ${TABLE()}`);
      const live = rows.filter((r) => s(r.obj["Coil No"]) !== "");
      const existing = Object.fromEntries((await sql`select coil_no, excel_hash, excel_row, dispatch_id, status, source, updated_at, mirrored_at from coils`).map((c) => [c.coil_no, c]));
      const fresh = [];

      // 1. import / update BSC-owned fields
      for (const r of live) {
        const o = r.obj, coil = s(o["Coil No"]), h = hash(o), ex = existing[coil];
        if (ex && ex.source === "transporter") { if (ex.excel_row !== r.i) await sql`update coils set excel_row=${r.i} where coil_no=${coil}`; continue; }
        const own = { entry_date: fromSerial(o["Entry Date"]), so_no: s(o["NMDC SO No"]), do_no: s(o["NMDC Invoice / DO No"]), form: s(o["Form"]), grade: s(o["Grade"]),
          thk: n(o["Thk (mm)"]), width: n(o["Width (mm)"]), length: n(o["Length (mm)"]), weight: n(o["Weight (MT)"]), destination: s(o["Destination"]),
          transporter_name: s(o["Transporter Name"]), transporter_mobile: s(o["Transporter Mobile"]).replace(/\D/g, ""), remarks: s(o["Remarks"]) };
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
          out.imported++; fresh.push({ coil, ...own });
        } else if (ex.excel_hash !== h || ex.excel_row !== r.i) {
          await sql`update coils set entry_date=${own.entry_date}, so_no=${own.so_no}, do_no=${own.do_no}, form=${own.form}, grade=${own.grade}, thk=${own.thk}, width=${own.width}, length=${own.length},
              weight=${own.weight}, destination=${own.destination}, transporter_name=${own.transporter_name}, transporter_mobile=${own.transporter_mobile}, remarks=${own.remarks},
              excel_row=${r.i}, excel_hash=${h}, transporter_id=${tid} where coil_no=${coil}`;
          if (ex.excel_hash !== h) out.updated++;
          if (!ex.dispatch_id && own.transporter_mobile) fresh.push({ coil, ...own });   // transporter added later
        }
      }

      // 2. dispatches + WhatsApp
      const groups = {};
      for (const f of fresh) { if (!f.transporter_mobile) continue; const k = f.do_no + "|" + f.transporter_mobile; (groups[k] ||= []).push(f); }
      for (const k of Object.keys(groups)) {
        const g = groups[k], f = g[0];
        let disp = (await sql`select * from dispatches where do_no=${f.do_no} and transporter_mobile=${f.transporter_mobile} order by id desc limit 1`)[0];
        if (!disp) {
          const token = crypto.randomBytes(9).toString("base64url");
          disp = (await sql`insert into dispatches (token, do_no, transporter_name, transporter_mobile, transporter_id) values (${token}, ${f.do_no}, ${f.transporter_name}, ${f.transporter_mobile}, ${tmap[f.transporter_mobile]?.id ?? null}) returning *`)[0];
          out.dispatches++;
        }
        await sql`update coils set dispatch_id=${disp.id} where coil_no = any(${g.map((x) => x.coil)})`;
        try {
          const count = (await sql`select count(*)::int as c from coils where dispatch_id=${disp.id}`)[0].c;
          const standing = tmap[disp.transporter_mobile];
          const r = await sendDispatchLink({ mobile: disp.transporter_mobile, name: disp.transporter_name, coils: count, doNo: disp.do_no, link: `${base}/t/${standing ? standing.token : disp.token}` });
          await sql`update dispatches set notified_at=now(), notify_error=${r.skipped || null} where id=${disp.id}`;
          if (!r.skipped) out.notified++;
        } catch (e) { out.errors.push(e.message); await sql`update dispatches set notify_error=${e.message} where id=${disp.id}`; }
      }

      // 3a. append transporter-created coils that are not yet in Excel
      const fullRow = (c) => headers.map((h) => ({ "Sl No": "", "Entry Date": serial(c.entry_date), "NMDC SO No": c.so_no || "", "NMDC Invoice / DO No": c.do_no || "", "Coil No": c.coil_no,
        "Form": c.form || "", "Grade": c.grade || "", "Thk (mm)": c.thk ?? "", "Width (mm)": c.width ?? "", "Length (mm)": c.length ?? "", "Weight (MT)": c.weight ?? "", "Status": c.status,
        "Vehicle No": c.vehicle_no || "", "Transporter Name": c.transporter_name || "", "Transporter Mobile": c.transporter_mobile || "", "LR No": c.lr_no || "",
        "Dispatch Date (ex NMDC)": serial(c.dispatch_date), "Expected Arrival": serial(c.expected_arrival), "Actual Arrival": serial(c.actual_arrival), "Destination": c.destination || "",
        "Remarks": c.remarks || "", "Entered By": c.entered_by || c.transporter_name || "", "Last Updated": "" }[h] ?? ""));
      const news = await sql`select * from coils where source='transporter' and excel_row is null order by created_at`;
      if (news.length) {
        const lay = await tableLayout(sess);
        const firstEmpty = live.length;                                   // next body row inside the (pre-sized) table
        for (const [k, c] of news.entries()) {
          const rowIdx = firstEmpty + k, sheetRow = lay.headerRow + 1 + rowIdx;
          if (rowIdx >= rows.length) { out.errors.push("tbl_Dispatch is full — extend the table"); break; }
          const vals = fullRow(c); vals[headers.indexOf("Sl No")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",ROW()-${lay.headerRow})`;
          vals[headers.indexOf("Last Updated")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",NOW())`;
          const fmt = headers.map((h) => (/Date|Arrival|Updated/.test(h) ? "dd-mmm-yy" : (/Weight/.test(h) ? "#,##0.000" : (/Mobile|No$|Coil No|LR No/.test(h) ? "@" : "General"))));
          await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${colLetter(0)}${sheetRow}:${colLetter(headers.length - 1)}${sheetRow}')`, { values: [vals], numberFormat: [fmt] });
          await sql`update coils set excel_row=${rowIdx}, mirrored_at=now() where coil_no=${c.coil_no}`;
          live.push({ i: rowIdx, obj: { "Coil No": c.coil_no } });
          out.appended++;
        }
      }
      // 3b. mirror changed fields back to Excel
      const dirty = await sql`select * from coils where excel_row is not null and (mirrored_at is null or updated_at > mirrored_at)`;
      if (dirty.length) {
        const lay = await tableLayout(sess);
        const firstCol = headers.indexOf("Status"), lastCol = headers.indexOf("Actual Arrival");
        const byRow = Object.fromEntries(live.map((r) => [r.i, r]));
        for (const c of dirty) {
          const r = byRow[c.excel_row];
          if (!r || s(r.obj["Coil No"]) !== c.coil_no) { out.errors.push(`row moved for ${c.coil_no}`); continue; }
          const sheetRow = lay.headerRow + 1 + c.excel_row;
          if (c.source === "transporter") {                                // transporter owns the whole row
            const vals = fullRow(c); vals[headers.indexOf("Sl No")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",ROW()-${lay.headerRow})`; vals[headers.indexOf("Last Updated")] = `=IF(${colLetter(headers.indexOf("Coil No"))}${sheetRow}="","",NOW())`;
            await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${colLetter(0)}${sheetRow}:${colLetter(headers.length - 1)}${sheetRow}')`, { values: [vals] });
            await sql`update coils set mirrored_at=now() where coil_no=${c.coil_no}`; out.mirrored++; continue;
          }
          const addr = `${colLetter(firstCol)}${sheetRow}:${colLetter(lastCol)}${sheetRow}`;
          const vals = headers.slice(firstCol, lastCol + 1).map((h) => ({
            "Status": c.status, "Vehicle No": c.vehicle_no || "", "Transporter Name": c.transporter_name || "", "Transporter Mobile": c.transporter_mobile || "",
            "LR No": c.lr_no || "", "Dispatch Date (ex NMDC)": serial(c.dispatch_date), "Expected Arrival": serial(c.expected_arrival), "Actual Arrival": serial(c.actual_arrival) }[h]));
          const fmt = headers.slice(firstCol, lastCol + 1).map((h) => (/Date|Arrival/.test(h) ? "dd-mmm-yy" : "@"));
          await sess.patch(`/worksheets('${encodeURIComponent(lay.sheet)}')/range(address='${addr}')`, { values: [vals], numberFormat: [fmt] });
          await sql`update coils set mirrored_at=now() where coil_no=${c.coil_no}`;
          out.mirrored++;
        }
      }
    });
    json(res, 200, { ok: true, ...out });
  } catch (e) { json(res, 500, { ok: false, ...out, error: e.message }); }
}
