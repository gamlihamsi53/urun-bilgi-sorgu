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
    const shareUrl = event.queryStringParameters && event.queryStringParameters.url;
    const mode = (event.queryStringParameters && event.queryStringParameters.mode) || "ping";

    if (!shareUrl) {
      return { statusCode: 400, body: JSON.stringify({ error: "Missing ?url=" }) };
    }

    const shareToken = toShareToken(shareUrl);
    const contentUrl = `https://api.onedrive.com/v1.0/shares/${shareToken}/root/content`;

    if (mode === "ping") {
      const r = await fetch(contentUrl, {
        method: "GET",
        redirect: "manual",
        headers: { "User-Agent": "Mozilla/5.0", Accept: "*/*" },
      });

      const location = r.headers.get("location");
      return {
        statusCode: 200,
        body: JSON.stringify({
          httpStatus: r.status,
          hasLocation: !!location,
          locationHost: location ? new URL(location).host : null,
          contentType: r.headers.get("content-type"),
          contentLength: r.headers.get("content-length"),
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

    return { statusCode: 400, body: JSON.stringify({ error: "Invalid mode. Use mode=ping" }) };
  } catch (e) {
    return { statusCode: 500, body: JSON.stringify({ error: String(e) }) };
  }
};
