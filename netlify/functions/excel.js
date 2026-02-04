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
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "Missing ?url=" }),
      };
    }

    const shareToken = toShareToken(shareUrl);
    const contentUrl =
