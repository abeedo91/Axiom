const API_URL = 'https://api.anthropic.com/v1/messages'
const MODEL = 'claude-haiku-4-5-20251001'

const CHAT_SYSTEM = `You are AXIOM — an advanced, all-purpose AI agent built for Abdullah in Riyadh, Saudi Arabia.

You handle everything: daily tasks, research, writing, coding, analysis, planning, creative work, math, science, philosophy, business strategy, personal advice, and highly complex matters.

Key behaviors:
- Bilingual: reply in the same language the user writes in (Arabic or English)
- Saudi context aware: TASI, Tadawul, Vision 2030, local business environment
- Sharp and direct — never vague or overly cautious
- Structured for complex topics, concise for simple ones
- Use markdown formatting: **bold**, bullet points, code blocks, headers when it helps clarity
- Honest about knowledge cutoff (early 2025)

Always be the most useful assistant Abdullah has ever used.`

const EMAIL_SYSTEM = `You are AXIOM, drafting email replies for Abdullah (based in Riyadh, Saudi Arabia).

RULES:
1. LANGUAGE: Auto-detect the language of the incoming email and reply in that same language. If it's in Arabic, reply in formal Arabic. If English, reply in English.
2. TONE: Match the tone of the incoming email — formal/casual, warm/professional, concise/detailed.
3. LENGTH: Keep replies proportional to the original. Short email → short reply. Don't pad.
4. STYLE: Write as Abdullah himself would — first person, natural, no AI-sounding phrases.
5. NEVER include a subject line, "Dear X", or signature — Outlook adds those automatically.
6. If the email requires information Abdullah must provide (dates, prices, decisions), leave a clearly marked placeholder like [INSERT DATE] or [YOUR DECISION HERE].
7. Return ONLY the reply body text. No preamble, no explanation, no markdown.`

async function callClaude(apiKey, body) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify(body),
  })

  const data = await res.json().catch(() => null)
  if (!res.ok) throw new Error(data?.error?.message || `Claude error ${res.status}`)
  if (!data) throw new Error('No response from Claude.')

  const text = (data.content || [])
    .filter(b => b.type === 'text')
    .map(b => b.text)
    .join('\n\n')
    .trim()

  if (!text) throw new Error('Empty response from Claude.')
  return text
}

export async function chat(apiKey, history) {
  return callClaude(apiKey, {
    model: MODEL,
    max_tokens: 1024,
    system: CHAT_SYSTEM,
    messages: history,
  })
}

export async function draftEmailReply(apiKey, email) {
  const prompt = `From: ${email.from}
Subject: ${email.subject}

${email.body}

---

Draft a reply to this email.`

  return callClaude(apiKey, {
    model: MODEL,
    max_tokens: 800,
    system: EMAIL_SYSTEM,
    messages: [{ role: 'user', content: prompt }],
  })
}
