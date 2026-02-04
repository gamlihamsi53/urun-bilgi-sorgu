// netlify/functions/excel.js

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

export async function handler(event) {
  try {
    const shareUrl = event.queryStringParameters?.url;
    const mode = event.queryStringParameters?.mode || "ping";
    if (!shareUrl) {
      return { statusCode: 400, body: JSON.stringify({ error: "Missing ?url=" }) };
    }

    const shareToken = toShareToken(shareUrl);
    const contentUrl = `https://api.onedrive.com/v1.0/shares/${shareToken}/root/content`;

    // ping: sadece response status/headers gör
    if (mode === "ping") {
      const r = await fetch(contentUrl, {
        method: "GET",
        redirect: "manual", // önemli: redirect takip ETME
        headers: { "User-Agent": "Mozilla/5.0", Accept: "*/*" },
      });

      // Not: manual redirect'te 302 bekleyebilirsin (bu iyi işaret!)
      const location = r.headers.get("location");
      const ct = r.headers.get("content-type");
      const cl = r.headers.get("content-length");

      return {
        statusCode: 200,
        body: JSON.stringify({
          ok: r.ok,
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

    // download: redirect takip et, dosyayı gerçekten indir, sadece boyut döndür
    if (mode === "download") {
      const r = await fetch(contentUrl, {
        method: "GET",
        redirect: "follow",
        headers: { "User-Agent": "Mozilla/5.0", Accept: "*/*" },
      });

      if (!r.ok) {
        return {
          statusCode: 200,
          body: JSON.stringify({
            ok: false,
            httpStatus: r.status,
            contentType: r.headers.get("content-type"),
            bodySnippet: await safeText(r),
          }),
        };
      }

      const buf = Buffer.from(await r.arrayBuffer());
      return {
        statusCode: 200,
        body: JSON.stringify({
          ok: true,
          httpStatus: r.status,
          contentType: r.headers.get("content-type"),
          bytes: buf.length,
          firstBytesHex: buf.slice(0, 12).toString("hex"), // hızlı sanity check
          note: "bytes > 0 ise indirme çalışıyor.",
        }),
      };
    }

    return { statusCode: 400, body: JSON.stringify({ error: "Invalid mode. Use ping|download" }) };
  } catch (e) {
    return { statusCode: 500, body: JSON.stringify({ error: String(e) }) };
  }
}
