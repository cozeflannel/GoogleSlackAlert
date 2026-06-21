import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

/**
 * Verifies a Slack request using HMAC-SHA256 signature.
 */
async function verifySlackSignature(req: Request, rawBody: string): Promise<boolean> {
  const signingSecret = Deno.env.get("SLACK_SIGNING_SECRET");
  if (!signingSecret) {
    console.error("SLACK_SIGNING_SECRET is not set");
    return false;
  }

  const timestamp = req.headers.get("X-Slack-Request-Timestamp");
  const slackSig = req.headers.get("X-Slack-Signature");
  if (!timestamp || !slackSig) return false;

  const nowSec = Math.floor(Date.now() / 1000);
  if (Math.abs(nowSec - parseInt(timestamp, 10)) > 300) return false;

  const sigBaseString = `v0:${timestamp}:${rawBody}`;
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(signingSecret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );

  const signatureBuffer = await crypto.subtle.sign("HMAC", key, encoder.encode(sigBaseString));
  const hexSignature =
    "v0=" +
    Array.from(new Uint8Array(signatureBuffer))
      .map((b) => b.toString(16).padStart(2, "0"))
      .join("");

  // NOTE: This can be logged if needed (but avoid logging secrets).
  return hexSignature === slackSig;
}

serve(async (req) => {
  try {
    const url = new URL(req.url);
    const action = url.searchParams.get("action");

    console.log("---- Incoming request ----");
    console.log("method:", req.method);
    console.log("path:", url.pathname);
    console.log("action:", action);
    console.log("content-type:", req.headers.get("content-type"));

    // 1. Handle GET requests
    if (req.method === "GET" || req.method === "HEAD") {
      if (action === "slack_oauth") {
        const spreadsheetId = url.searchParams.get("state");
        const clientId = Deno.env.get("SLACK_CLIENT_ID");
        const redirectUri = `${Deno.env.get("SUP_URL")}/functions/v1/bot?action=oauth_callback`;
        const scopes =
          "chat:write,chat:write.public,channels:read,channels:join,app_mentions:read,im:write";

        const slackAuthUrl =
          "https://slack.com/oauth/v2/authorize" +
          "?client_id=" +
          encodeURIComponent(clientId!) +
          "&scope=" +
          encodeURIComponent(scopes) +
          "&redirect_uri=" +
          encodeURIComponent(redirectUri) +
          "&state=" +
          encodeURIComponent(spreadsheetId!);

        console.log("Redirecting to Slack OAuth URL");
        return Response.redirect(slackAuthUrl, 302);
      }

      if (action === "oauth_callback") {
        console.log("OAuth callback route hit (TODO not implemented).");
        return new Response("OAuth Callback Received. Please close this window.", {
          headers: { "Content-Type": "text/html" },
        });
      }

      if (action === "privacy") {
        console.log("Privacy route hit.");
        return new Response(`<html><body><h1>Privacy Policy</h1></body></html>`, {
          headers: { "Content-Type": "text/html" },
        });
      }

      console.log("Default GET route hit.");
      return new Response("SheetAlerts Edge Function Active", { status: 200 });
    }

    // 2. Handle POST requests
    const bodyText = await req.text();
    console.log("raw body length:", bodyText.length);

    const verified = await verifySlackSignature(req, bodyText);
    console.log("slack signature verified:", verified);

    if (!verified) {
      console.log("Unauthorized: signature verification failed.");
      return new Response("Unauthorized", { status: 401 });
    }

    const contentType = req.headers.get("content-type") ?? "";
    console.log("Parsing request based on content-type:", contentType);

    let payload: any;
    if (contentType.includes("application/x-www-form-urlencoded")) {
      const form = new URLSearchParams(bodyText);
      const raw = form.get("payload");
      payload = raw ? JSON.parse(raw) : {};
    } else {
      try {
        payload = JSON.parse(bodyText);
      } catch {
        payload = {};
      }
    }

    console.log("payload keys:", payload ? Object.keys(payload) : []);
    console.log("payload.type:", payload?.type);

    // Slack Handshake
    if (payload.type === "url_verification") {
      console.log("Handling Slack url_verification handshake.");
      const challenge = payload.challenge;

      const resPayload = { challenge };
      console.log("returning JSON:", resPayload);

      return new Response(JSON.stringify(resPayload), {
        headers: { "Content-Type": "application/json" },
      });
    }

    // TODO: Implement block_actions and view_submission handlers
    console.log("No specific handler yet; returning ok:true");
    const resPayload = { ok: true };
    console.log("returning JSON:", resPayload);

    return new Response(JSON.stringify(resPayload), {
      headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("Edge Function Error:", msg);
    return new Response(`Internal Server Error: ${msg}`, { status: 500 });
  }
});