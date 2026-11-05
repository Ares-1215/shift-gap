// Edge Function: shiftgap — 班次空檔分析網站（ares-1215.github.io/shift-gap）的共用檔案庫
// 全部用 POST JSON：{ action, code, ... }
//   login                                  看板碼或管理碼 → { role: 'view' | 'admin' }
//   list                                   （看板碼）所有檔案的 metadata（不含內容）
//   get    { key }                         （看板碼）單一檔案，內容 base64
//   put    { kind, day, name, data, summary, uploader, mime }   （管理碼）新增／覆蓋，key = kind/day
//   delete { key }                         （管理碼）
// 通行碼放在資料表 sg_settings（view_code／admin_code／retain_days）；sg_files 只有 service role 能讀寫。
// 部署：tools\deploy.ps1（verify_jwt=false，這裡自己驗通行碼）
import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const SB_URL = Deno.env.get("SUPABASE_URL")!;
const SB_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const MAX_B64 = 12 * 1024 * 1024; // 單檔上限（base64 後）

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "content-type, authorization, apikey, x-client-info",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Content-Type": "application/json; charset=utf-8",
};
const J = (o: unknown, status = 200) => new Response(JSON.stringify(o), { status, headers: cors });
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

function rest(path: string, init: RequestInit = {}) {
  return fetch(`${SB_URL}/rest/v1/${path}`, {
    ...init,
    headers: { apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`, "Content-Type": "application/json", ...(init.headers ?? {}) },
  });
}
async function must(r: Response, what: string) {
  if (!r.ok) throw new Error(`${what} ${r.status}: ${await r.text()}`);
  return r;
}
async function settings(): Promise<Record<string, string>> {
  const r = await must(await rest("sg_settings?select=key,value"), "settings");
  const rows = (await r.json()) as { key: string; value: string }[];
  return Object.fromEntries(rows.map((x) => [x.key, x.value]));
}

const KEY_RE = /^(schedule|needs)\/[0-9A-Za-z_-]{1,24}$/;
const DAY_RE = /^[0-9A-Za-z_-]{1,24}$/;
const META = "key,kind,day,name,size,mime,summary,uploader,uploaded_at";

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  if (req.method !== "POST") return J({ error: "method not allowed" }, 405);
  try {
    const body = await req.json();
    const action = String(body.action ?? "");
    const code = String(body.code ?? "").trim();
    const S = await settings();
    const isAdmin = !!code && code === (S.admin_code ?? "");
    const isView = isAdmin || (!!code && code === (S.view_code ?? ""));
    const retain = Math.max(1, Number(S.retain_days ?? 21) || 21);

    if (action === "login") {
      if (!isView) { await sleep(700); return J({ error: "通行碼錯誤" }, 401); }
      return J({ role: isAdmin ? "admin" : "view", retain_days: retain });
    }
    if (!isView) return J({ error: "需要通行碼" }, 401);

    if (action === "list") {
      const r = await must(await rest(`sg_files?select=${META}&order=kind.asc,day.asc`), "list");
      return J({ files: await r.json(), now: new Date().toISOString(), retain_days: retain });
    }
    if (action === "get") {
      const key = String(body.key ?? "");
      if (!KEY_RE.test(key)) return J({ error: "key 格式錯誤" }, 400);
      const r = await must(await rest(`sg_files?key=eq.${encodeURIComponent(key)}&select=${META},data`), "get");
      const rows = await r.json();
      if (!rows.length) return J({ error: "找不到檔案" }, 404);
      return J(rows[0]);
    }
    if (!isAdmin) return J({ error: "需要管理碼" }, 403);

    if (action === "put") {
      const kind = String(body.kind ?? "");
      const day = String(body.day ?? "").trim();
      if (kind !== "schedule" && kind !== "needs") return J({ error: "kind 只能是 schedule 或 needs" }, 400);
      if (!DAY_RE.test(day)) return J({ error: "day 格式錯誤（MMDD 或英數）" }, 400);
      const data = String(body.data ?? "");
      if (!data || data.length > MAX_B64 || !/^[A-Za-z0-9+/=\r\n]+$/.test(data.slice(0, 4000))) {
        return J({ error: "data 要是 base64，且不超過上限" }, 400);
      }
      const row = {
        key: `${kind}/${day}`,
        kind,
        day,
        name: String(body.name ?? `${day}.bin`).slice(0, 200),
        size: Math.floor(data.replace(/[\r\n=]/g, "").length * 3 / 4),
        mime: String(body.mime ?? "").slice(0, 100) || null,
        data,
        summary: body.summary && typeof body.summary === "object" ? body.summary : null,
        uploader: String(body.uploader ?? "").slice(0, 40) || null,
        uploaded_at: new Date().toISOString(),
      };
      await must(await rest("sg_files?on_conflict=key", {
        method: "POST",
        headers: { Prefer: "resolution=merge-duplicates,return=minimal" },
        body: JSON.stringify([row]),
      }), "put");
      // 清掉過期檔案（預設 21 天）
      const cutoff = new Date(Date.now() - retain * 86400000).toISOString();
      await rest(`sg_files?uploaded_at=lt.${encodeURIComponent(cutoff)}`, { method: "DELETE", headers: { Prefer: "return=minimal" } });
      return J({ ok: true, key: row.key, uploaded_at: row.uploaded_at, size: row.size });
    }
    if (action === "delete") {
      const key = String(body.key ?? "");
      if (!KEY_RE.test(key)) return J({ error: "key 格式錯誤" }, 400);
      await must(await rest(`sg_files?key=eq.${encodeURIComponent(key)}`, { method: "DELETE", headers: { Prefer: "return=minimal" } }), "delete");
      return J({ ok: true });
    }
    return J({ error: "unknown action" }, 400);
  } catch (e) {
    return J({ error: String((e as Error).message ?? e) }, 500);
  }
});
