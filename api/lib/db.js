import { neon } from "@neondatabase/serverless";
export const sql = neon(process.env.DATABASE_URL);
export const STATUSES = ["Under Loading", "Loaded", "In-Transit", "Arrived"];
export function json(res, status, body) { res.setHeader("Cache-Control", "no-store"); res.status(status).json(body); }
export const d = (v) => (v ? String(v).slice(0, 10) : null);

/** CORS for the dashboard hosted on GitHub Pages. ALLOWED_ORIGINS = comma-separated list, e.g. https://dispatch.bharatsteels.in */
export function cors(req, res) {
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
  if (req.method === "OPTIONS") { res.status(204).end(); return true; }
  return false;
}
