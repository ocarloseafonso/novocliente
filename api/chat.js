// api/chat.js — Vercel Serverless Function
// A chave da OpenAI fica aqui no servidor, NUNCA exposta no navegador.

export default async function handler(req, res) {
  // Só aceita POST
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido' });
  }

  // Pega a chave das variáveis de ambiente da Vercel (segura)
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'Chave da OpenAI não configurada no servidor.' });
  }

  try {
    const { messages, model = 'gpt-4o-mini', temperature = 0.85, max_tokens = 800 } = req.body;

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({ model, messages, temperature, max_tokens })
    });

    if (!response.ok) {
      const errData = await response.json().catch(() => ({}));
      return res.status(response.status).json({ error: errData });
    }

    const data = await response.json();
    return res.status(200).json(data);

  } catch (err) {
    console.error('Erro no proxy OpenAI:', err);
    return res.status(500).json({ error: err.message });
  }
}
