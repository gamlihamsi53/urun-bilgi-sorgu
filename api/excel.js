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
        "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7"
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

module.exports = async (req, res) => {
  try {
    const mode = (req.query.mode || "resolve").toString();
    const shareUrl = req.query.url ? req.query.url.toString() : null;

    if (!shareUrl) {
      res.status(400).json({ error: "Missing ?url=" });
      return;
    }

    if (mode === "resolve") {
      const resolved = await resolveRedirectChain(shareUrl, 12);
      res.status(200).json(resolved);
      return;
    }

    if (mode === "download" || mode === "parse") {
      const resolved = await resolveRedirectChain(shareUrl, 12);

      // login'e düşüyorsa burada dur
      let host = null;
      try { host = new URL(resolved.finalUrl).host; } catch {}
      if (host && host.includes("login.live.com")) {
        res.status(200).json({
          ok: false,
          error: "Login required (link anonymous değil veya erişim kısıtlı).",
          resolved
        });
        return;
      }

      const r = await fetch(resolved.finalUrl, {
        method: "GET",
        redirect: "follow",
        headers: {
          "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
          "Accept": "*/*"
        },
      });

      if (!r.ok) {
        res.status(200).json({
          step: "download",
          ok: false,
          httpStatus: r.status,
          contentType: r.headers.get("content-type"),
          bodySnippet: await safeText(r),
          resolved
        });
        return;
      }

      const buf = Buffer.from(await r.arrayBuffer());

      if (mode === "download") {
        res.status(200).json({
          step: "download",
          ok: true,
          httpStatus: r.status,
          contentType: r.headers.get("content-type"),
          bytes: buf.length,
          firstBytesHex: buf.slice(0, 16).toString("hex"),
          resolved
        });
        return;
      }

      // parse
      const wb = XLSX.read(buf, { type: "buffer" });
      const firstSheetName = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

      res.status(200).json({
        step: "parse",
        ok: true,
        sheet: firstSheetName,
        rowCount: rows.length,
        rows
      });
      return;
    }

    res.status(400).json({ error: "Invalid mode. Use resolve|download|parse" });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
};
