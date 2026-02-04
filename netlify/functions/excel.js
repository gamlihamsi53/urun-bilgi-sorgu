const XLSX = require("xlsx");

// ✅ Buraya OneDrive paylaşım linkini (1drv.ms) koy
const SHARE_URL = "https://1drv.ms/x/c/5e38cf4df60786e7/IQDRH_yMYQbeS7QHKUC0YdnrAaG1m5KBMun6eBzFMrsotRU?e=7a1Jst";

const SHEET_NAME = "Mal Tanımı";
const COL_KEY = "Ürün Açıklaması";

const FIELDS = [
  "KDV Oranı",
  "Birim",
  "Alış Fiyatı (KDV Hariç)",
  "Son Satın Alma Tarihi",
  "Liste Fiyatı (KDV Hariç)",
  "Son Liste Fiyat Güncelleme Tarihi",
  "Marj",
  "Stok"
];

function normalizeTR(s) {
  return (s ?? "")
    .toString()
    .trim()
    .toLocaleLowerCase("tr-TR")
    .replace(/\s+/g, " ");
}

function absUrl(next, base) {
  try { return new URL(next, base).toString(); } catch { return next; }
}

function pickParam(urlStr, key) {
  try {
    const u = new URL(urlStr);
    return u.searchParams.get(key);
  } catch {
    return null;
  }
}

function buildDownloadUrl(resid, authkey) {
  return `https://onedrive.live.com/download?resid=${encodeURIComponent(resid)}&authkey=${encodeURIComponent(authkey)}`;
}

async function resolveToDownloadUrl(shareUrl, debug) {
  let cur = shareUrl;

  for (let i = 0; i < 12; i++) {
    debug.step = `redirect_${i}`;
    const resp = await fetch(cur, { redirect: "manual" });

    if ([301, 302, 303, 307, 308].includes(resp.status)) {
      const loc = resp.headers.get("location");
      if (!loc) break;
      cur = absUrl(loc, cur);
      continue;
    }

    // Redirect bitince URL’de resid/authkey yakalamayı dene
    const residFromUrl = pickParam(cur, "resid");
    const authFromUrl = pickParam(cur, "authkey");
    if (residFromUrl && authFromUrl) {
      debug.resolvedBy = "url_params";
      return buildDownloadUrl(residFromUrl, authFromUrl);
    }

    // HTML geldiyse içinden resid/authkey yakalamayı dene
    const ct = resp.headers.get("content-type") || "";
    debug.contentType = ct;

    if (ct.includes("text/html")) {
      const html = await resp.text();

      const residMatch = html.match(/resid=([^&"'<\s]+)/i);
      const authMatch = html.match(/authkey=([^&"'<\s]+)/i);
      if (residMatch && authMatch) {
        debug.resolvedBy = "html_regex";
        return buildDownloadUrl(residMatch[1], authMatch[1]);
      }

      const dlMatch = html.match(/https:\/\/onedrive\.live\.com\/download[^"'<\s]+/i);
      if (dlMatch) {
        debug.resolvedBy = "html_download_url";
        return dlMatch[0];
      }
    }

    break;
  }

  throw new Error("Redirect zincirinden indirilebilir (resid/authkey) link üretemedim.");
}

exports.handler = async () => {
  const debug = {
    step: "start",
    shareUrlHost: null,
    contentType: null,
    resolvedBy: null
  };

  try {
    debug.shareUrlHost = (() => {
      try { return new URL(SHARE_URL).host; } catch { return "invalid-url"; }
    })();

    const headers = {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "public, max-age=600, stale-while-revalidate=86400"
    };

    debug.step = "resolve_download_url";
    const downloadUrl = await resolveToDownloadUrl(SHARE_URL, debug);

    debug.step = "download_binary";
    const fileResp = await fetch(downloadUrl, { redirect: "follow" });
    if (!fileResp.ok) throw new Error(`download HTTP ${fileResp.status}`);

    const buf = await fileResp.arrayBuffer();

    debug.step = "xlsx_parse";
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[SHEET_NAME];
    if (!ws) throw new Error(`Sheet bulunamadı: "${SHEET_NAME}" (mevcut: ${wb.SheetNames.join(", ")})`);

    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    debug.step = "build_map";
    const items = [];
    const map = {};

    for (const r of rows) {
      const key = r[COL_KEY];
      if (!key) continue;

      const name = key.toString();
      const norm = normalizeTR(name);

      const payload = { [COL_KEY]: name };
      for (const f of FIELDS) payload[f] = r[f] ?? "";

      items.push(name);
      map[norm] = payload;
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        updatedAt: new Date().toISOString(),
        sheet: SHEET_NAME,
        keyColumn: COL_KEY,
        fields: FIELDS,
        items,
        map
      })
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: e?.message || String(e), debug })
    };
  }
};
