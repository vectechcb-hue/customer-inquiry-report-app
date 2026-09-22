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
  const key = await crypto.subtle.importKey(
    "raw",
    enc.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(body));
  const expected = btoa(String.fromCharCode(...new Uint8Array(sig)));
  return expected === signature;
}

function isLikelyChannelSecret(value) {
  return /^[0-9a-fA-F]{32}$/.test(String(value || "").trim());
}

function stringBindings(env) {
  const out = {};
  for (const key of Object.keys(env || {})) {
    const value = env[key];
    if (typeof value === "string" && value.trim()) out[key] = value.trim();
  }
  return out;
}

async function discoverLineCredentials(env) {
  const bindings = stringBindings(env);
  const namedSecret = bindings.LINE_CHANNEL_SECRET || "";
  const namedAccess = bindings.LINE_CHANNEL_ACCESS_TOKEN || bindings.LINE_CHANNEL_ACCESS_TC || "";

  let channelSecret = namedSecret;
  let secretSource = namedSecret ? "LINE_CHANNEL_SECRET" : "";

  if (!isLikelyChannelSecret(channelSecret)) {
    const candidate = Object.entries(bindings).find(([key, value]) =>
      key !== "LINE_READ_API_KEY" &&
      isLikelyChannelSecret(value)
    );
    if (candidate) {
      channelSecret = candidate[1];
      secretSource = candidate[0];
    }
  }

  let accessToken = namedAccess;
  let accessSource = namedAccess ? (bindings.LINE_CHANNEL_ACCESS_TOKEN ? "LINE_CHANNEL_ACCESS_TOKEN" : "LINE_CHANNEL_ACCESS_TC") : "";

  // If the token variable name is wrong, discover a working LINE Bot token
  // without exposing any token value.
  if (accessToken && (await getBotInfo(accessToken)).ok) {
    return { channelSecret, accessToken, secretSource, accessSource };
  }

  const candidates = Object.entries(bindings).filter(([key, value]) =>
    key !== "LINE_CHANNEL_SECRET" &&
    key !== "LINE_READ_API_KEY" &&
    key !== secretSource &&
    value.length >= 20 &&
    value !== channelSecret
  );

  for (const [key, value] of candidates) {
    const info = await getBotInfo(value);
    if (info.ok) {
      accessToken = value;
      accessSource = key;
      break;
    }
  }

  return { channelSecret, accessToken, secretSource, accessSource };
}

async function getProfile(userId, token) {
  if (!userId || !token) return "";
  try {
    const r = await fetch(
      "https://api.line.me/v2/bot/profile/" + encodeURIComponent(userId),
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!r.ok) return "";
    const j = await r.json();
    return j.displayName || "";
  } catch (_) {
    return "";
  }
}

async function ensureDiagnosticsTable(env) {
  await env.DB.prepare(`
    CREATE TABLE IF NOT EXISTS webhook_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      timestamp INTEGER NOT NULL,
      event_count INTEGER NOT NULL,
      signature_present INTEGER NOT NULL,
      signature_valid INTEGER NOT NULL,
      event_types TEXT NOT NULL,
      processed INTEGER NOT NULL DEFAULT 0
    )
  `).run();
}

async function logWebhook(env, data) {
  try {
    await ensureDiagnosticsTable(env);
    const result = await env.DB.prepare(
      "INSERT INTO webhook_logs (timestamp,event_count,signature_present,signature_valid,event_types,processed) VALUES (?,?,?,?,?,?)"
    ).bind(
      Number(data.timestamp || Date.now()),
      Number(data.eventCount || 0),
      data.signaturePresent ? 1 : 0,
      data.signatureValid ? 1 : 0,
      String(data.eventTypes || ""),
      data.processed ? 1 : 0
    ).run();
    return result.meta?.last_row_id || null;
  } catch (_) {
    return null;
  }
}

async function markWebhookProcessed(env, id) {
  if (!id) return;
  try {
    await env.DB.prepare("UPDATE webhook_logs SET processed=1 WHERE id=?").bind(id).run();
  } catch (_) {}
}

async function processEvents(events, env) {
  let processed = 0;
  for (const ev of events) {
    if (!ev || !ev.type) continue;

    const eventId =
      ev.webhookEventId ||
      (ev.timestamp + "-" + (ev.message?.id || ev.type || "") + "-" + (ev.source?.userId || ""));

    const userId = ev.source?.userId || "";
    const messageType = ev.type === "message"
      ? (ev.message?.type || "")
      : ev.type;

    const messageText =
      ev.type === "message" && ev.message?.type === "text"
        ? (ev.message.text || "")
        : "";

    const creds = await discoverLineCredentials(env);
    const displayName =
      (ev.source?.type === "user")
        ? await getProfile(userId, creds.accessToken)
        : "";

    await env.DB.prepare(
      "INSERT OR IGNORE INTO line_events (event_id,user_id,timestamp,message_type,message_text,display_name,salesperson,raw_json) VALUES (?,?,?,?,?,?,?,?)"
    ).bind(
      eventId,
      userId,
      Number(ev.timestamp || Date.now()),
      messageType,
      messageText,
      displayName,
      "",
      JSON.stringify(ev)
    ).run();

    processed++;
  }
  return processed;
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response("", { status: 204, headers: jsonHeaders });
    }

    const url = new URL(request.url);

    if (request.method === "POST" && url.pathname === "/webhook") {
      const rawBody = await request.text();

      let payload;
      try {
        payload = JSON.parse(rawBody);
      } catch (_) {
        return response({ ok: false, error: "invalid json" }, 400);
      }

      const events = Array.isArray(payload.events) ? payload.events : [];
      const signature = request.headers.get("x-line-signature") || "";

      // LINE Developers Verify sends events: [] and requires an HTTP 200 response.
      if (events.length === 0) {
        await logWebhook(env, {
          timestamp: Date.now(),
          eventCount: 0,
          signaturePresent: !!signature,
          signatureValid: true,
          eventTypes: "",
          processed: true
        });
        return response({ ok: true, verify: true });
      }

      const credentials = await discoverLineCredentials(env);
      const signatureValid = await verifySignature(credentials.channelSecret, rawBody, signature);
      const eventTypes = events.map(e => e?.type || "").filter(Boolean).join(",");

      const logId = await logWebhook(env, {
        timestamp: Date.now(),
        eventCount: events.length,
        signaturePresent: !!signature,
        signatureValid,
        eventTypes: eventTypes + ":" + (credentials.secretSource || "none") + ":" + (credentials.accessSource || "none"),
        processed: false
      });

      if (!signatureValid) {
        return response({ ok: false, error: "invalid signature" }, 401);
      }

      try {
        await processEvents(events, env);
        await markWebhookProcessed(env, logId);
        return response({ ok: true });
      } catch (err) {
        return response({
          ok: false,
          error: "database processing failed",
          detail: String(err?.message || err || "unknown")
        }, 500);
      }
    }

    if (request.method === "GET" && url.pathname === "/messages") {
      const readKey = request.headers.get("X-API-Key") || "";
      if (!env.LINE_READ_API_KEY || readKey !== env.LINE_READ_API_KEY) {
        return response({ ok: false, error: "unauthorized" }, 401);
      }

      const from = Date.parse(
        url.searchParams.get("from") || "1970-01-01T00:00:00.000Z"
      );
      const to = Date.parse(
        url.searchParams.get("to") || "2999-12-31T00:00:00.000Z"
      );

      if (!Number.isFinite(from) || !Number.isFinite(to)) {
        return response({ ok: false, error: "invalid date range" }, 400);
      }

      const result = await env.DB.prepare(
        "SELECT event_id,user_id,timestamp,message_type,message_text,display_name,salesperson FROM line_events WHERE timestamp >= ? AND timestamp < ? ORDER BY timestamp DESC"
      ).bind(from, to).all();

      const events = (result.results || []).map(x => ({
        eventId: x.event_id,
        userId: x.user_id,
        timestamp: new Date(Number(x.timestamp)).toISOString(),
        messageType: x.message_type,
        text: x.message_text || "",
        displayName: x.display_name || "",
        salesperson: x.salesperson || ""
      }));

      return response({ ok: true, events });
    }

    if (request.method === "GET" && url.pathname === "/status") {
      try {
        await ensureDiagnosticsTable(env);
        const eventCount = await env.DB.prepare(
          "SELECT COUNT(*) AS count FROM line_events"
        ).first();
        const webhookLogCount = await env.DB.prepare(
          "SELECT COUNT(*) AS count FROM webhook_logs"
        ).first();
        const lastWebhook = await env.DB.prepare(
          "SELECT timestamp,event_count,signature_present,signature_valid,event_types,processed FROM webhook_logs ORDER BY id DESC LIMIT 1"
        ).first();
        const lastEvent = await env.DB.prepare(
          "SELECT timestamp,message_type FROM line_events ORDER BY timestamp DESC LIMIT 1"
        ).first();

        const credentials = await discoverLineCredentials(env);
        const [secretHash, accessHash, botInfo] = await Promise.all([
          sha256Hex(credentials.channelSecret),
          sha256Hex(credentials.accessToken),
          getBotInfo(credentials.accessToken)
        ]);

        return response({
          ok: true,
          service: "vectech-line-customer-api",
          credentialSource: {
            secret: credentials.secretSource || "none",
            accessToken: credentials.accessSource || "none"
          },
          channelSecretLooksValid: isLikelyChannelSecret(credentials.channelSecret),
          channelSecretFingerprint: secretHash ? secretHash.slice(0, 12) : "",
          accessTokenLength: credentials.accessToken.length,
          accessTokenFingerprint: accessHash ? accessHash.slice(0, 12) : "",
          accessTokenBotInfo: {
            ok: botInfo.ok,
            status: botInfo.status,
            basicId: botInfo.basicId || "",
            userId: botInfo.userId || ""
          },
          eventCount: Number(eventCount?.count || 0),
          webhookLogCount: Number(webhookLogCount?.count || 0),
          lastWebhook: lastWebhook ? {
            timestamp: new Date(Number(lastWebhook.timestamp)).toISOString(),
            eventCount: Number(lastWebhook.event_count || 0),
            signaturePresent: !!lastWebhook.signature_present,
            signatureValid: !!lastWebhook.signature_valid,
            eventTypes: lastWebhook.event_types || "",
            processed: !!lastWebhook.processed
          } : null,
          lastEvent: lastEvent ? {
            timestamp: new Date(Number(lastEvent.timestamp)).toISOString(),
            messageType: lastEvent.message_type || ""
          } : null
        });
      } catch (err) {
        return response({ ok: false, error: "status query failed", detail: String(err?.message || err || "unknown") }, 500);
      }
    }

    if (request.method === "GET" && url.pathname === "/diagnostics") {
      const readKey = request.headers.get("X-API-Key") || "";
      if (!env.LINE_READ_API_KEY || readKey !== env.LINE_READ_API_KEY) {
        return response({ ok: false, error: "unauthorized" }, 401);
      }

      await ensureDiagnosticsTable(env);

      const eventCount = await env.DB.prepare(
        "SELECT COUNT(*) AS count FROM line_events"
      ).first();

      const webhookLogs = await env.DB.prepare(
        "SELECT id,timestamp,event_count,signature_present,signature_valid,event_types,processed FROM webhook_logs ORDER BY id DESC LIMIT 20"
      ).all();

      return response({
        ok: true,
        eventCount: Number(eventCount?.count || 0),
        webhookLogs: webhookLogs.results || []
      });
    }

    if (request.method === "GET" && url.pathname === "/health") {
      return response({ ok: true, service: "vectech-line-customer-api" });
    }

    return response({ ok: false, error: "not found" }, 404);
  }
};
