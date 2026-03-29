// AralMate Cloudflare Worker — Kimi API Proxy
// Deploy this to Cloudflare Workers.
// Set KIMI_API_KEY and ARALMATE_SECRET as encrypted environment variables (Secrets).
//
// IMPORTANT: Use the endpoint that matches your Moonshot AI account platform:
//   api.moonshot.ai  — international accounts registered at platform.moonshot.ai
//   api.moonshot.cn  — China/mainland accounts registered at platform.moonshot.cn
// Keys are NOT interchangeable between platforms — wrong endpoint = 401 error.
const KIMI_ENDPOINT = "https://api.moonshot.ai/v1/chat/completions";

export default {
  async fetch(request, env) {
    // 1. Handle CORS preflight
    if (request.method === "OPTIONS") {
      return corsResponse(null, 204);
    }

    // 2. Only allow POST
    if (request.method !== "POST") {
      return corsResponse(JSON.stringify({ error: "Method not allowed" }), 405);
    }

    // 3. Validate secret header
    const secret = request.headers.get("X-AralMate-Secret");
    if (!secret || secret !== env.ARALMATE_SECRET) {
      return corsResponse(JSON.stringify({ error: "Unauthorized" }), 401);
    }

    // 4. Parse request body
    let body;
    try {
      body = await request.json();
    } catch {
      return corsResponse(JSON.stringify({ error: "Invalid JSON" }), 400);
    }

    // 5. Forward to Kimi API
    const kimiResponse = await fetch(KIMI_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${env.KIMI_API_KEY}`
      },
      body: JSON.stringify(body)
    });

    const data = await kimiResponse.json();
    return corsResponse(JSON.stringify(data), kimiResponse.status);
  }
};

function corsResponse(body, status = 200) {
  return new Response(body, {
    status,
    headers: {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, X-AralMate-Secret"
    }
  });
}
