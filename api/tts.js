// api/tts.js — Vercel Serverless Function para TTS (Text-to-Speech)
// A chave da OpenAI fica aqui no servidor, NUNCA exposta no navegador.

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido' });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'Chave da OpenAI não configurada no servidor.' });
  }

  try {
    const { input, voice = 'nova', model = 'tts-1' } = req.body;

    const response = await fetch('https://api.openai.com/v1/audio/speech', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({ model, voice, input })
    });

    if (!response.ok) {
      const errData = await response.json().catch(() => ({}));
      return res.status(response.status).json({ error: errData });
    }

    // A resposta é um áudio binário (buffer)
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // Enviar o buffer como resposta mp3
    res.setHeader('Content-Type', 'audio/mpeg');
    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate'); // Cache do áudio para economizar créditos
    return res.status(200).send(buffer);

  } catch (err) {
    console.error('Erro no proxy TTS OpenAI:', err);
    return res.status(500).json({ error: err.message });
  }
}
