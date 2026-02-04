// netlify/functions/excel.js
// mode=resolve  -> 1drv.ms redirect zincirini çözer, final URL'i verir
// mode=download -> çözdüğü final URL üzerinden indirir, byte sayısını döndürür

async function safeText(resp, n = 300) {
  try {
    const t = await resp.text();
    return t.slice(0, n);
  } catch {
    return null;
  }
}

function isRedirect(status) {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}

async function resolveRedirectChain(startUrl, maxHops = 10) {
  let url = startUrl;
  const visited = [];

  for (let i = 0; i < maxHops; i++) {
    visited.push(url);

    const r = await fetch(url, {
      method: "GET",
      redirect: "manual",
      headers: {
        // Tarayıcı gibi görünmek 1drv.ms tarafında kritik olabiliyor
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept":
          "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
      },
    });

    // redirect değilse burada dururuz
    if (!isRedirect(r.status)) {
      return {
        ok: r.ok,
        finalStatus: r.status,
        finalUrl: url,
        visited,
        contentType: r.headers.get("content-type"),
        bodySnippet: r.ok ? null : await safeText(r),
      };
    }

    const loc = r.headers.get("location");
    if (!loc) {
      return {
        ok: false,
        finalStatus: r.status,
        finalUrl: url,
        visited,
        error: "Redirect var ama Location header yok",
      };
    }

    // relative olabilir
    url = new URL(loc, url).toString();
  }

  return {
    ok: false,
    finalStatus: 0,
    finalUrl: url,
    visited,
    error: `Max hops (${maxHops}) aşıldı`,
  };
}

exports.handler = async function (event) {
  try {
    const qs = event.queryStringParameters || {};
    const mode = qs.mode || "resolve";
    const shareUrl = qs.url;

    if (!shareUrl) {
      return {
        statusCode: 400,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({ error: "Missing ?url=" }),
      };
    }

    if (mode === "resolve") {
      const result = await resolveRedirectChain(shareUrl, 12);
      return {
        statusCode: 200,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify(result),
      };
    }

    if (mode === "download") {
      // Önce çöz
      const resolved = await resolveRedirectChain(shareUrl, 12);

      // Eğer çözme aşamasında redirect olmayan 4xx/5xx aldıysak, indirmenin anlamı yok
      if (!resolved || (!resolved.ok && resolved.finalStatus >= 400)) {
        return {
          statusCode: 200,
          headers: { "content-type": "application/json; charset=utf-8" },
          body: JSON.stringify({ step: "resolve", resolved }),
        };
      }

      // Çözülen URL ile indir
      const r = await fetch(resolved.finalUrl, {
        method: "GET",
        redirect: "follow",
        headers: {
          "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
          Accept: "*/*",
        },
      });

      if (!r.ok) {
        return {
          statusCode: 200,
          headers: { "content-type": "application/json; charset=utf-8" },
          body: JSON.stringify({
            step: "download",
            ok: false,
            httpStatus: r.status,
            contentType: r.headers.get("content-type"),
            bodySnippet: await safeText(r),
            resolved,
          }),
        };
      }

      const buf = Buffer.from(await r.arrayBuffer());
      return {
        statusCode: 200,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({
          step: "download",
          ok: true,
          httpStatus: r.status,
          contentType: r.headers.get("content-type"),
          bytes: buf.length,
          firstBytesHex: buf.slice(0, 16).toString("hex"),
          resolved,
        }),
      };
    }

    return {
      statusCode: 400,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: "Invalid mode. Use resolve|download" }),
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: String(e) }),
    };
  }
};
