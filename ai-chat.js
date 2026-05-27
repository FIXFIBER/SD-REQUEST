export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { provider, messages } = req.body;

  try {
    if (provider === 'cohere') {
      const response = await fetch("https://api.cohere.com/v2/chat", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${process.env.COHERE_API_KEY}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: "command-r-plus-08-2024",
          messages,
          temperature: 0.5,
          max_tokens: 1200
        })
      });
      const data = await response.json();
      return res.status(response.status).json(data);
    }

    if (provider === 'openrouter') {
      const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${process.env.OPENROUTER_API_KEY}`,
          "Content-Type": "application/json",
          "X-Title": "FiberOne Admin"
        },
        body: JSON.stringify({
          model: "google/gemini-2.0-flash-exp:free",
          messages,
          temperature: 0.5,
          max_tokens: 1200
        })
      });
      const data = await response.json();
      return res.status(response.status).json(data);
    }

    return res.status(400).json({ error: 'Invalid provider specified' });
  } catch (error) {
    return res.status(500).json({ error: 'Internal Server Error', details: error.message });
  }
}