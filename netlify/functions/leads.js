const { getStore } = require("@netlify/blobs");

// centralized lead storage using Netlify Blobs — this is server-side and
// shared across everyone who uses the app, unlike localStorage which was
// stuck per-browser. writes (saving a lead) are open to anyone using the
// app, since that's the whole point of "save lead". reads (viewing/
// exporting everything saved) require the admin key — set ADMIN_KEY in
// Netlify env vars, then visit your site as:
//   https://your-site.netlify.app/?admin=<that key>

exports.handler = async (event) => {
  const store = getStore("plum-personal-leads");

  if (event.httpMethod === "POST") {
    try {
      const lead = JSON.parse(event.body);
      const key = `lead-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      await store.setJSON(key, lead);
      return { statusCode: 200, body: JSON.stringify({ saved: true }) };
    } catch (err) {
      return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (event.httpMethod === "GET") {
    const providedKey = event.queryStringParameters?.key;
    const adminKey = process.env.ADMIN_KEY;

    if (!adminKey) {
      return { statusCode: 500, body: JSON.stringify({ error: "ADMIN_KEY is not set on this Netlify site" }) };
    }
    if (providedKey !== adminKey) {
      return { statusCode: 401, body: JSON.stringify({ error: "unauthorized" }) };
    }

    try {
      const { blobs } = await store.list();
      const leads = await Promise.all(blobs.map((b) => store.get(b.key, { type: "json" })));
      return { statusCode: 200, body: JSON.stringify({ leads: leads.filter(Boolean) }) };
    } catch (err) {
      return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
    }
  }

  return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
};
