import { serve } from "https://deno.land/std@0.168.0/http/server.ts"

async function openModalAsync(payload: any, action: any) {
  try {
    console.log("openModalAsync started. trigger_id present?", Boolean(payload?.trigger_id));
    console.log("views.open trigger_id:", payload?.trigger_id);

    const actionValue = JSON.parse(action.value);
    const { spreadsheetId, rowIndex } = actionValue;

    const dbUrl = `${Deno.env.get("SUPABASE_URL")}/rest/v1/bot_configs?spreadsheet_id=eq.${spreadsheetId}&select=*`;
    const dbRes = await fetch(dbUrl, {
      headers: {
        'apikey': Deno.env.get("BOT_SECRET_KEY")!,
        'Authorization': `Bearer ${Deno.env.get("BOT_SECRET_KEY")}`
      }
    });

    const dbData = await dbRes.json();
    if (!dbData?.length) throw new Error("No config found — re-save settings in Google Sheet");

    const { token, headers, actionable_cols } = dbData[0];

    const blocks: any[] = [
      { type: 'section', text: { type: 'mrkdwn', text: `Updating row *${rowIndex}*.` } }
    ];

    (actionable_cols || []).forEach((colIdx: string) => {
      const headerName = (headers || [])[parseInt(colIdx)] || `Column ${colIdx}`;
      blocks.push({
        type: 'input', block_id: `col_${colIdx}_block`,
        element: { type: 'plain_text_input', action_id: `col_${colIdx}_input` },
        label: { type: 'plain_text', text: headerName }
      });
    });

    blocks.push({
      type: 'input', block_id: 'notes_block', optional: true,
      element: { type: 'plain_text_input', action_id: 'notes_input', multiline: true },
      label: { type: 'plain_text', text: 'Notes' }
    });

    const slackRes = await fetch('https://slack.com/api/views.open', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        trigger_id: payload.trigger_id,
        view: {
          type: 'modal', callback_id: 'resolve_alert_modal',
          private_metadata: action.value,
          title: { type: 'plain_text', text: 'Resolve Alert' },
          submit: { type: 'plain_text', text: 'Submit' },
          close: { type: 'plain_text', text: 'Cancel' },
          blocks
        }
      })
    });

    const result = await slackRes.json();
    console.log("views.open result:", result);
    if (!result.ok) console.error("views.open failed:", result.error);
    else console.log("views.open succeeded for row", rowIndex);
  } catch (err) {
    console.error("openModalAsync error:", err.message);
  }
}

serve(async (req) => {
  try {
    const rawUrl = Deno.env.get("GAS_WEBAPP_URL")?.trim();
    if (!rawUrl) {
      console.error("ERROR: GAS_WEBAPP_URL secret is empty.");
      return new Response("Missing Configuration", { status: 500 });
    }

    const url = new URL(req.url);
    const gasUrl = new URL(rawUrl);
    url.searchParams.forEach((val, key) => gasUrl.searchParams.set(key, val));

    // 1. Handle GET requests
    if (req.method === 'GET' || req.method === 'HEAD') {
      const action = url.searchParams.get("action");

      if (action === 'diag') {
        try {
          const dbUrl = `${Deno.env.get("SUPABASE_URL")}/rest/v1/bot_configs?select=*`;
          const dbRes = await fetch(dbUrl, {
            headers: {
              'apikey': Deno.env.get("BOT_SECRET_KEY")!,
              'Authorization': `Bearer ${Deno.env.get("BOT_SECRET_KEY")}`
            }
          });
          const data = await dbRes.json();
          return new Response(JSON.stringify({
            status: 'OK',
            db_response: dbRes.status,
            count: Array.isArray(data) ? data.length : 0,
            data: data
          }), { headers: { "Content-Type": "application/json" } });
        } catch (err) {
          return new Response(JSON.stringify({ status: 'ERROR', message: err.message }), {
            status: 500,
            headers: { "Content-Type": "application/json" }
          });
        }
      }

      if (action === "privacy") {
        return new Response(`
          <html><body style="font-family:sans-serif; padding:40px; line-height:1.6; max-width:800px; margin:auto;">
            <h1>Privacy Policy - SheetAlerts</h1>
            <p>SheetAlerts only accesses your spreadsheets to monitor for specific condition changes you define.</p>
            <p>We do not store your spreadsheet data on our servers. All alerts are forwarded directly to Slack/Gmail.</p>
            <p>For support, please contact: support@example.com</p>
          </body></html>
        `, { headers: { "Content-Type": "text/html" } });
      }

      if (!action) {
        return new Response(`
          <html><body style="font-family:sans-serif; text-align:center; padding:100px;">
            <h1 style="color:#0052cc;">SheetAlerts is Active</h1>
            <p style="font-size:1.2em;">Your intelligent spreadsheet-to-Slack bridge is running.</p>
            <p style="color:#666;">Status: Connected to Google Apps Script</p>
          </body></html>
        `, { headers: { "Content-Type": "text/html" } });
      }

      console.log("PROXY: Forwarding GET to Google...");
      return await fetch(gasUrl.toString(), { redirect: 'follow' });
    }

    // 2. Handle POST requests
    const contentType = req.headers.get("content-type") ?? "";
    const bodyText = await req.text();
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

    // Instant Slack Handshake
    if (payload.type === 'url_verification') {
      console.log("Slack handshake detected.");
      return new Response(
        JSON.stringify({ challenge: payload.challenge }),
        { headers: { "Content-Type": "application/json" } }
      );
    }

    // Intercept "Take Action" Modal — respond immediately, open modal async
    if (payload.type === 'block_actions') {
      const action = payload.actions?.[0];
      if (action?.action_id === 'open_action_modal') {
        console.log("open_action_modal intercepted, trigger_id:", payload.trigger_id);
        await openModalAsync(payload, action);
        return new Response(JSON.stringify({ ok: true }), {
          headers: { "Content-Type": "application/json" }
        });
      }
    }

    // Forward all other POSTs to Google
    console.log("PROXY: Forwarding POST to Google...");
    const fwdHeaders = new Headers();
    if (req.headers.get("content-type")) {
      fwdHeaders.set("content-type", req.headers.get("content-type")!);
    }
    const response = await fetch(gasUrl.toString(), {
      method: 'POST',
      headers: fwdHeaders,
      body: bodyText,
      redirect: 'follow'
    });
    return response;

  } catch (err) {
    console.error("PROXY CRASH:", err.message);
    return new Response(`Internal Server Error: ${err.message}`, { status: 500 });
  }
});
ASH:", err.message);
    return new Response(`Internal Server Error: ${err.message}`, { status: 500 });
  }
});
