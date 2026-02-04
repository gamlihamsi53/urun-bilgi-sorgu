const XLSX = require("xlsx");

// ✅ Buraya OneDrive paylaşım linkini (1drv.ms) koy
const SHARE_URL = "PASTE_YOUR_1DRV_MS_LINK_HERE";

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

// OneDrive share URL -> shareId (base64url)
function toShareId(url) {
  const b64 = Buffer.from(url, "utf8").toString("base64");
  const b64url = b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return "u!" + b64url;
}

exports.handler = async () => {
  const debug = {
    step: "start",
    shareUrlHost: null,
    contentType: null
  };

  try {
    if (!SHARE_URL || SHARE_URL.includes("PASTE_")) {
      return {
        statusCode: 500,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({ error: "SHARE_URL ayarlanmadı (1drv.ms linkini yapıştır).", debug })
      };
    }

    debug.shareUrlHost = (() => {
      try { return new URL(SHARE_URL).host; } catch { return "invalid-url"; }
    })();

    // Netlify cache
    const headers = {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "public, max-age=600, stale-while-revalidate=86400"
    };

    // 1) ShareId oluştur
    debug.step = "shareId";
    const shareId = toShareId(SHARE_URL);

    // 2) “shares” endpoint’inden driveItem bilgisi al (downloadUrl almak için)
    // Bu endpoint public share linklerle çalışır.
    debug.step = "shares_api";
    const metaUrl = `https://api.onedrive.com/v1.0/shares/${shareId}/driveItem?$select=id,name,@microsoft.graph.downloadUrl`;

    const metaResp = await fetch(metaUrl, { redirect: "follow" });
    if (!metaResp.ok) {
      throw new Error(`shares meta HTTP ${metaResp.status}`);
    }
    const meta = await metaResp.json();

    const downloadUrl = meta?.["@microsoft.graph.downloadUrl"];
    if (!downloadUrl) {
      throw new Error("downloadUrl bulunamadı (paylaşım linki view değil mi?).");
    }

    // 3) Dosyayı gerçek downloadUrl’den indir
    debug.step = "download_binary";
    const fileResp = await fetch(downloadUrl, { redirect: "follow" });
    if (!fileResp.ok) {
      throw new Error(`download HTTP ${fileResp.status}`);
    }
    debug.contentType = fileResp.headers.get("content-type") || "";
    const buf = await fileResp.arrayBuffer();

    // 4) XLSX parse
    debug.step = "xlsx_parse";
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[SHEET_NAME];
    if (!ws) throw new Error(`Sheet bulunamadı: "${SHEET_NAME}" (mevcut: ${wb.SheetNames.join(", ")})`);

    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    // 5) Autocomplete + map
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

    debug.step = "done";

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
      body: JSON.stringify({
        error: e?.message || String(e),
        debug
      })
    };
  }
};
