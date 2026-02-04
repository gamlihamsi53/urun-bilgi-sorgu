// netlify/functions/excel.js
// MODE:
//  - ping: OneDrive share linkine erişim var mı (redirect geliyor mu) test eder
//
// Çağrı örneği:
// /.netlify/functions/excel?mode=ping&url=<1drv.ms link>

function toShareToken(shareUrl) {
  const b64 = Buffer.from(shareUrl, "utf8").toString("base64");
  const b64url = b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return "u!" + b64url;
}

async function safeText(resp, n = 300) {
  try {
    const t = await resp.text();
    return t.slice(0, n);
  } catch {
    return null;
  }
}

exports.handler = async function (event) {
  try {
    const qs = event.queryStringParameters || {};
    const shareUrl = qs.url;
    const mode = qs.mode || "ping";

    if (!shareUrl) {
      return {
        statusCode: 400,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({ error: "Missing ?url=" }),
      };
    }

    const shareToken = toShareToken(shareUrl);
    const contentUrl = `https://api.onedrive.com/v1.0/shares/${shareToken}/root/content`;

    if (mode === "ping") {
      const r = await fetch(contentUrl, {
        method: "GET",
        redirect: "manual", // redirect'i takip ETME
        headers: {
          "User-Agent": "Mozilla/5.0",
          Accept: "*/*",
        },
      });

      const location = r.headers.get("location");
      const ct = r.headers.get("content-type");
      const cl = r.headers.get("content-length");

      return {
        statusCode: 200,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({
          httpStatus: r.status,
          hasLocation: !!location,
          locationHost: location ? new URL(location).host : null,
          contentType: ct,
          contentLength: cl,
          note:
            r.status >= 300 && r.status < 400
              ? "302/301 aldıysan iyi: redirect var, erişim var."
              : r.ok
              ? "200 aldıysan çok iyi: direkt içerik dönmüş."
              : "4xx/5xx ise izin/link tipi sorunu olabilir.",
          bodySnippet: r.ok ? null : await safeText(r),
        }),
      };
    }

    return {
      statusCode: 400,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: "Invalid mode. Use mode=ping" }),
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: String(e) }),
    };
  }
};
