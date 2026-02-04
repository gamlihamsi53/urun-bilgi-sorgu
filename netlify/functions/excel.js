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

  for (let i = 0; i < 15; i++) {
    debug.step = `redirect_${i}`;
    const resp = await fetch(cur, { redirect: "manual" });

    // 1) Header redirect
    if ([301, 302, 303, 307, 308].includes(resp.status)) {
      const loc = resp.headers.get("location");
      if (!loc) throw new Error(`Redirect HTTP ${resp.status} ama Location yok`);
      cur = absUrl(loc, cur);
      continue;
    }

    // Redirect değilse status kontrol
    if (!resp.ok) {
      throw new Error(`Redirect olmayan cevap: HTTP ${resp.status}`);
    }

    // 2) URL üzerinde resid/authkey var mı?
    const residFromUrl = pickParam(cur, "resid");
    const authFromUrl = pickParam(cur, "authkey");
    if (residFromUrl && authFromUrl) {
      debug.resolvedBy = "url_params";
      return buildDownloadUrl(residFromUrl, authFromUrl);
    }

    // 3) HTML/boş content-type durumlarında body’yi incele (OneDrive bazen burada JS/meta refresh veriyor)
    const ct = (resp.headers.get("content-type") || "").toLowerCase();
    debug.contentType = ct;

    const text = await resp.text();

    // 3a) meta refresh ile yönlendirme
    const meta = text.match(/http-equiv=["']refresh["'][^>]*content=["'][^"']*url=([^"']+)["']/i);
    if (meta?.[1]) {
      cur = absUrl(meta[1], cur);
      continue;
    }

    // 3b) JS yönlendirme (window.location / location.href)
    const js = text.match(/(?:window\.location|location\.href)\s*=\s*["']([^"']+)["']/i);
    if (js?.[1]) {
      cur = absUrl(js[1], cur);
      continue;
    }

    // 3c) Sayfa içinde onedrive.live.com linki ara
    const urls = [...text.matchAll(/https?:\/\/[^"'<> \n]+/g)].map(m => m[0]);

    // Önce doğrudan download linki yakala
    const directDl = urls.find(u => /onedrive\.live\.com\/download/i.test(u));
    if (directDl) {
      debug.resolvedBy = "html_download_url";
      return directDl;
    }

    // Sonra embed linkinden resid/authkey yakala
    const embed = urls.find(u => /onedrive\.live\.com\/embed/i.test(u) && /resid=/i.test(u) && /authkey=/i.test(u));
    if (embed) {
      debug.resolvedBy = "html_embed_url";
      const resid = pickParam(embed, "resid");
      const auth = pickParam(embed, "authkey");
      if (resid && auth) return buildDownloadUrl(resid, auth);
    }

    // Son olarak onedrive.live.com’a giden herhangi bir linki takip et
    const anyLive = urls.find(u => /onedrive\.live\.com/i.test(u));
    if (anyLive) {
      cur = anyLive;
      continue;
    }

    throw new Error("HTML içinde yönlendirme/download linki bulunamadı.");
  }

  throw new Error("Redirect zinciri çok uzadı, indirilebilir link üretilemedi.");
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
