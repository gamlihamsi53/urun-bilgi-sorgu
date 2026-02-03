const XLSX = require("xlsx");

const ONEDRIVE_XLSX_URL =
  "https://onedrive.live.com/download?resid=5E38CF4DF60786E7!8341D4774B0845D1987763213E7BE629&authkey=!ABCDEF123456";

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
    if (!ONEDRIVE_XLSX_URL) {
      return {
        statusCode: 500,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({ error: "ONEDRIVE_XLSX_URL boş." })
      };
    }

    // Netlify cache: 10 dk taze, 1 gün stale-while-revalidate
    const headers = {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "public, max-age=600, stale-while-revalidate=86400"
    };

    const resp = await fetch(ONEDRIVE_XLSX_URL, { redirect: "follow" });
    if (!resp.ok) {
      throw new Error("OneDrive HTTP " + resp.status);
    }

    const contentType = resp.headers.get("content-type") || "";
    const buf = await resp.arrayBuffer();

    // OneDrive bazen HTML sayfa döndürebiliyor (indirilemediyse).
    // XLSX okumadan önce hızlı kontrol:
    if (contentType.includes("text/html")) {
      throw new Error("OneDrive HTML döndürdü (download link doğru mu?).");
    }

    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[SHEET_NAME];
    if (!ws) throw new Error(`Sheet bulunamadı: "${SHEET_NAME}"`);

    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

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
      body: JSON.stringify({ error: e?.message || String(e) })
    };
  }
};
