export default async function handler(req, res) {
    // Solo permitir POST
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const GROQ_API_KEY = process.env.GROQ_API_KEY;

    if (!GROQ_API_KEY) {
        return res.status(500).json({ error: 'API Key no configurada en el servidor' });
    }

    try {
        const { messages, model, temperature } = req.body;

        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${GROQ_API_KEY}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                model: model || "llama-3.1-8b-instant",
                messages: messages,
                temperature: temperature || 0.1
            })
        });

        const data = await response.json();

        if (!response.ok) {
            return res.status(response.status).json({ error: data.error || 'Error en GROQ API' });
        }

        return res.status(200).json(data);

    } catch (error) {
        console.error('Error en proxy GROQ:', error);
        return res.status(500).json({ error: 'Error interno del servidor' });
    }
}
