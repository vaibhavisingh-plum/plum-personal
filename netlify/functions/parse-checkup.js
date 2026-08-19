// Netlify serverless function — keeps the Anthropic API key on the server,
// never exposed to the browser. The React app calls this at
// /.netlify/functions/parse-checkup instead of calling Claude directly.
// Accepts either an image (jpg/png/etc) or a PDF of a health checkup report.
//
// SETUP REQUIRED (see instructions from Claude):
// 1. In Netlify site settings > Environment variables, add:
//      ANTHROPIC_API_KEY = <your key from console.anthropic.com>
// 2. This only works when deployed via git or `netlify deploy` — a plain
//    drag-and-drop of the `dist` folder does NOT include functions.

const KNOWN_CONDITION_IDS = ["cataract", "hypertension", "thyroid", "diabetes", "cardiac", "cancer"];

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const { imageBase64, mediaType } = JSON.parse(event.body);

    if (!imageBase64 || !mediaType) {
      return { statusCode: 400, body: JSON.stringify({ error: "Missing imageBase64 or mediaType" }) };
    }

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      return {
        statusCode: 500,
        body: JSON.stringify({ error: "ANTHROPIC_API_KEY is not set on this Netlify site" }),
      };
    }

    const isPdf = mediaType === "application/pdf";
    const fileBlock = isPdf
      ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: imageBase64 } }
      : { type: "image", source: { type: "base64", media_type: mediaType, data: imageBase64 } };

    const prompt = `Extract from this health checkup report:
1) the checkup date (format YYYY-MM-DD, or null if not clearly stated)
2) any medical conditions mentioned that match ONLY these ids: ${KNOWN_CONDITION_IDS.join(", ")}
3) a one-sentence plain summary of anything else notable (or empty string)

Respond with ONLY valid JSON, no markdown fences, no explanation:
{"checkupDate": "YYYY-MM-DD or null", "conditions": ["id1","id2"], "notes": "..."}`;

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 500,
        messages: [
          {
            role: "user",
            content: [fileBlock, { type: "text", text: prompt }],
          },
        ],
      }),
    });

    const data = await response.json();

    if (!response.ok) {
      return { statusCode: response.status, body: JSON.stringify({ error: data.error?.message || "Anthropic API error" }) };
    }

    const textBlock = data.content?.find((b) => b.type === "text");
    const raw = (textBlock?.text || "{}").replace(/```json|```/g, "").trim();

    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch {
      return { statusCode: 502, body: JSON.stringify({ error: "Could not parse model response", raw }) };
    }

    // guard: only keep condition ids we actually recognize in the app
    parsed.conditions = (parsed.conditions || []).filter((c) => KNOWN_CONDITION_IDS.includes(c));

    return { statusCode: 200, body: JSON.stringify(parsed) };
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
