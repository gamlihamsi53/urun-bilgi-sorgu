const XLSX = require("xlsx");

const ONEDRIVE_XLSX_URL = "https://1drv.ms/x/c/5e38cf4df60786e7/IQDRH_yMYQbeS7QHKUC0YdnrAaG1m5KBMun6eBzFMrsotRU?e=igHqiY"; // <- OneDrive VIEW linkini buraya yapıştır

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

exports.handler = async () => {
  try {
    if (!ONEDRIVE_XLSX_URL || ONEDRIVE_XLSX_URL.includes("PASTE_")) {
      return {
        statusCode: 500,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({ error: "ONEDRIVE_XLSX_URL ayarlanmadı (OneDrive VIEW linkini yapıştır)." })
      };
    }

    // Netlify cache: 10 dk taze, 1 gün stale-while-revalidate
    const headers = {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "public, max-age=600, stale-while-revalidate=86400"
    };

    const candidates = [
      ONEDRIVE_XLSX_URL,
      ONEDRIVE_XLSX_URL + (ONEDRIVE_XLSX_URL.includes("?") ? "&" : "?") + "download=1"
    ];

    let arrayBuffer = null;
    let lastErr = null;

    for (const url of candidates) {
      try {
        const resp = await fetch(url);
        if (!resp.ok) throw new Error("HTTP " + resp.status);
        arrayBuffer = await resp.arrayBuffer();
        break;
      } catch (e) {
        lastErr = e;
      }
    }

    if (!arrayBuffer) throw lastErr || new Error("Dosya indirilemedi.");

    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const ws = wb.Sheets[SHEET_NAME];
    if (!ws) throw new Error(`Sheet bulunamadı: "${SHEET_NAME}"`);

    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    // Autocomplete listesi + map
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
        items, // listbox için
        map    // arama için
      })
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: e?.message || String(e) })
    };
  }
};
