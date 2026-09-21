// WATI templates (Utility category). Variables are named.
//  nmdc_dispatch_link : name, coils, do_no, link
//  nmdc_transporter_link : name, link
async function send(mobile, template, parameters) {
  const base = process.env.WATI_API_URL, key = process.env.WATI_API_KEY;
  if (!base || !key) return { skipped: "WATI not configured" };
  const num = String(mobile).replace(/\D/g, "").replace(/^0+/, "");
  const wa = num.length === 10 ? "91" + num : num;
  const r = await fetch(`${base.replace(/\/$/, "")}/api/v1/sendTemplateMessage?whatsappNumber=${wa}`, {
    method: "POST", headers: { Authorization: "Bearer " + key, "Content-Type": "application/json" },
    body: JSON.stringify({ template_name: template, broadcast_name: template, parameters }),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || j.result === false) throw new Error("WATI: " + (j.info || j.message || r.status));
  return { ok: true };
}
export const sendDispatchLink = ({ mobile, name, coils, doNo, link }) =>
  send(mobile, process.env.WATI_TEMPLATE || "nmdc_dispatch_link", [{ name: "name", value: name || "Partner" }, { name: "coils", value: String(coils) }, { name: "do_no", value: doNo || "-" }, { name: "link", value: link }]);
export const sendTransporterLink = ({ mobile, name, link }) =>
  send(mobile, process.env.WATI_TEMPLATE_TRANSPORTER || "nmdc_transporter_link", [{ name: "name", value: name || "Partner" }, { name: "link", value: link }]);
