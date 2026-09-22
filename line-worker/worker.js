const jsonHeaders = {
  "Content-Type": "application/json; charset=utf-8",
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type, X-API-Key",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS"
};

function response(body, status = 200, extra = {}) {
  return new Response(JSON.stringify(body), { status, headers: { ...jsonHeaders, ...extra } });
}

async function verifySignature(secret, body, signature) {
  if (!secret || !signature) return false;
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey("raw", enc.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(body));
  const expected = btoa(String.fromCharCode(...new Uint8Array(sig)));
  return expected === signature;
}

async function getProfile(userId, token) {
  if (!userId || !token) return "";
  try {
    const r = await fetch("https://api.line.me/v2/bot/profile/" + encodeURIComponent(userId), { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) return "";
    const j = await r.json();
    return j.displayName || "";
  } catch (_) { return ""; }
}

async function processEvents(events, env) {
  for (const ev of events) {
    if (ev?.type !== "message") continue;
    const eventId = ev.webhookEventId || (ev.timestamp + "-" + (ev.message?.id || "") + "-" + (ev.source?.userId || ""));
    const userId = ev.source?.userId || "";
    const messageType = ev.message?.type || "";
    const messageText = messageType === "text" ? (ev.message?.text || "") : "";
    const displayName = await getProfile(userId, env.LINE_CHANNEL_ACCESS_TOKEN);
    await env.DB.prepare("INSERT OR IGNORE INTO line_events (event_id,user_id,timestamp,message_type,message_text,display_name,salesperson,raw_json) VALUES (?,?,?,?,?,?,?,?)").bind(eventId, userId, Number(ev.timestamp || Date.now()), messageType, messageText, displayName, "", JSON.stringify(ev)).run();
  }
}

export default {
  async fetch(request, env, ctx) {
    if (request.method === "OPTIONS") return new Response("", { status: 204, headers: jsonHeaders });
    const url = new URL(request.url);

    if (request.method === "POST" && url.pathname === "/webhook") {
      const rawBody = await request.text();
      let payload;
      try { payload = JSON.parse(rawBody); } catch (_) { return response({ ok: false, error: "invalid json" }, 400); }

      // LINE Developers "Verify" sends a connectivity-check POST with events: [].
      // Per LINE's documentation, this verification request should return HTTP 200.
      if (Array.isArray(payload.events) && payload.events.length === 0) {
        return response({ ok: true, verify: true });
      }

      const signature = request.headers.get("x-line-signature") || "";
      if (!(await verifySignature(env.LINE_CHANNEL_SECRET, rawBody, signature))) {
        return response({ ok: false, error: "invalid signature" }, 401);
      }

      ctx.waitUntil(processEvents(payload.events || [], env));
      return response({ ok: true });
    }

    if (request.method === "GET" && url.pathname === "/messages") {
      const readKey = request.headers.get("X-API-Key") || "";
      if (!env.LINE_READ_API_KEY || readKey !== env.LINE_READ_API_KEY) return response({ ok: false, error: "unauthorized" }, 401);
      const from = Date.parse(url.searchParams.get("from") || "1970-01-01T00:00:00.000Z");
      const to = Date.parse(url.searchParams.get("to") || "2999-12-31T00:00:00.000Z");
      if (!Number.isFinite(from) || !Number.isFinite(to)) return response({ ok: false, error: "invalid date range" }, 400);
      const result = await env.DB.prepare("SELECT event_id,user_id,timestamp,message_type,message_text,display_name,salesperson FROM line_events WHERE timestamp >= ? AND timestamp < ? ORDER BY timestamp DESC").bind(from, to).all();
      const events = (result.results || []).map(x => ({ eventId: x.event_id, userId: x.user_id, timestamp: new Date(Number(x.timestamp)).toISOString(), messageType: x.message_type, text: x.message_text || "", displayName: x.display_name || "", salesperson: x.salesperson || "" }));
      return response({ ok: true, events });
    }

    if (request.method === "GET" && url.pathname === "/health") return response({ ok: true, service: "vectech-line-customer-api" });
    return response({ ok: false, error: "not found" }, 404);
  }
};
