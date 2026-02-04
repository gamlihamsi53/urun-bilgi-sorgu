// netlify/functions/excel.js
// MODES:
//  - resolve : 1drv.ms redirect zincirini çözer (login mi, public mi?)
//  - download: indirip byte sayısını döndürür (xlsx mi geliyor?)
//  - parse   : indirip ilk sheet'i JSON döndürür (xlsx paketi gerekir)

const XLSX = require("xlsx");

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

async function resolveRedirectChain(startUrl, maxHops = 12) {
  let url = startUrl;
  const visited = [];

  for (let i = 0; i < maxHops; i++) {
    visited.push(url);

    const r = await fetch(url, {
      method: "GET",
      redirect: "manual",
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept":
          "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
      },
    });

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

    // 1) resolve
    if (mode === "resolve") {
      const resolved = await resolveRedirectChain(shareUrl, 12);
      return {
        statusCode: 200,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify(resolved),
      };
    }

    // 2) download
    if (mode === "download" || mode === "parse") {
      const resolved = await resolveRedirectChain(shareUrl, 12);

      // login'e düştüyse burada durdur
      const host = (() => {
        try {
          return new URL(resolved.finalUrl).host;
        } catch {
          return null;
        }
      })();

      if (host && host.includes("login.live.com")) {
        return {
          statusCode: 200,
          headers: { "content-type": "application/json; charset=utf-8" },
          body: JSON.stringify({
            ok: false,
            error: "Login required (link anonymous değil veya erişim kısıtlı).",
            resolved,
          }),
        };
      }

      // indir
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

      // sadece download sonucu
      if (mode === "download") {
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

      // 3) parse
      const wb = XLSX.read(buf, { type: "buffer" });
      const firstSheetName = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

      return {
        statusCode: 200,
        headers: { "content-type": "application/json; charset=utf-8" },
        body: JSON.stringify({
          step: "parse",
          ok: true,
          sheet: firstSheetName,
          rowCount: rows.length,
          rows,
        }),
      };
    }

    return {
      statusCode: 400,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: "Invalid mode. Use resolve|download|parse" }),
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: { "content-type": "application/json; charset=utf-8" },
      body: JSON.stringify({ error: String(e) }),
    };
  }
};
