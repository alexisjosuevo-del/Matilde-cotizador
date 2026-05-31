module.exports = async function handler(req, res) {
    // CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const GROQ_API_KEY = process.env.GROQ_API_KEY;

    if (!GROQ_API_KEY) {
        console.error('GROQ_API_KEY no está configurada');
        return res.status(500).json({ error: 'API Key no configurada en el servidor' });
    }

    try {
        const { messages, model, temperature } = req.body;

        if (!messages || !Array.isArray(messages)) {
            return res.status(400).json({ error: 'Se requiere el campo "messages" como array' });
        }

        const groqResponse = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${GROQ_API_KEY}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                model: model || "llama-3.1-8b-instant",
                messages: messages,
                temperature: temperature !== undefined ? temperature : 0.1
            })
        });

        const data = await groqResponse.json();

        if (!groqResponse.ok) {
            console.error('Error de GROQ:', JSON.stringify(data));
            return res.status(groqResponse.status).json(data);
        }

        return res.status(200).json(data);

    } catch (error) {
        console.error('Error en proxy GROQ:', error.message);
        return res.status(500).json({ error: 'Error interno: ' + error.message });
    }
};
