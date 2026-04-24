import { useState, useRef, useEffect, useCallback } from 'react'
import { chat, draftEmailReply } from './claude.js'
import {
  initMsal, signIn, signOut, getAccount,
  listEmails, getEmailBody, createDraftReply, stripHtml,
} from './outlook.js'

// ═══════════════════════════════════════════════════════════════
// Helpers
// ═══════════════════════════════════════════════════════════════

function parseMarkdown(text) {
  if (!text) return ''
  return text
    .replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) =>
      `<pre class="ax-code"><code>${code.trim().replace(/</g, '&lt;').replace(/>/g, '&gt;')}</code></pre>`
    )
    .replace(/`([^`\n]+)`/g, '<code class="ax-inline">$1</code>')
    .replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\*([^*\n]+)\*/g, '<em>$1</em>')
    .replace(/^### (.+)$/gm, '<h3 class="ax-h3">$1</h3>')
    .replace(/^## (.+)$/gm, '<h2 class="ax-h2">$1</h2>')
    .replace(/^# (.+)$/gm, '<h1 class="ax-h1">$1</h1>')
    .replace(/^[-•] (.+)$/gm, '<li class="ax-li">$1</li>')
    .replace(/^\d+\. (.+)$/gm, '<li class="ax-li-n">$1</li>')
    .replace(/\n\n/g, '<div class="ax-gap"></div>')
    .replace(/\n/g, '<br/>')
}

function formatRelative(iso) {
  const d = new Date(iso)
  const diff = (Date.now() - d.getTime()) / 1000
  if (diff < 60) return 'just now'
  if (diff < 3600) return `${Math.floor(diff / 60)}m ago`
  if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`
  if (diff < 604800) return `${Math.floor(diff / 86400)}d ago`
  return d.toLocaleDateString()
}

const SUGGESTED = [
  { icon: '📈', text: 'Analyze a TASI stock' },
  { icon: '✍️', text: 'Draft a professional email' },
  { icon: '💻', text: 'Help me write a script' },
  { icon: '🧠', text: 'Explain something complex' },
  { icon: '📊', text: 'Pros and cons of my idea' },
  { icon: '🗓️', text: 'Help me plan my week' },
]

// ═══════════════════════════════════════════════════════════════
// Styles (as JS objects, applied inline = guaranteed to render)
// ═══════════════════════════════════════════════════════════════

const S = {
  app: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    maxWidth: 720,
    margin: '0 auto',
    position: 'relative',
    zIndex: 1,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '14px 18px',
    paddingTop: 'calc(14px + env(safe-area-inset-top, 0px))',
    borderBottom: '1px solid rgba(255,255,255,0.06)',
    background: 'rgba(6,6,8,0.7)',
    backdropFilter: 'blur(20px) saturate(180%)',
    WebkitBackdropFilter: 'blur(20px) saturate(180%)',
    flexShrink: 0,
  },
  headerLeft: { display: 'flex', alignItems: 'center', gap: 12 },
  logo: {
    width: 36, height: 36, borderRadius: 11,
    background: 'linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%)',
    display: 'flex', alignItems: 'center', justifyContent: 'center',
    fontFamily: "'Syne', sans-serif",
    fontWeight: 800, fontSize: 11, color: 'white', letterSpacing: 0.5,
    boxShadow: '0 4px 14px rgba(124,109,250,0.4), inset 0 1px 0 rgba(255,255,255,0.2)',
    flexShrink: 0,
  },
  brand: {
    fontFamily: "'Syne', sans-serif",
    fontWeight: 700, fontSize: 17, letterSpacing: '0.06em',
    color: '#ffffff',
  },
  brandSub: {
    fontSize: 10, color: '#5a5a70',
    textTransform: 'uppercase', letterSpacing: '0.12em',
    fontWeight: 500, marginTop: 1,
  },
  iconBtn: {
    background: 'rgba(255,255,255,0.04)',
    border: '1px solid rgba(255,255,255,0.08)',
    borderRadius: 10,
    width: 36, height: 36,
    fontSize: 16, color: '#a0a0b8',
    cursor: 'pointer',
    display: 'flex', alignItems: 'center', justifyContent: 'center',
  },
  tabs: {
    display: 'flex',
    padding: '10px 14px 0',
    gap: 6,
    flexShrink: 0,
  },
  tabBtn: (active) => ({
    flex: 1,
    padding: '11px 12px',
    background: active ? 'rgba(124,109,250,0.18)' : 'rgba(255,255,255,0.03)',
    border: active ? '1px solid rgba(124,109,250,0.35)' : '1px solid rgba(255,255,255,0.06)',
    borderRadius: 11,
    color: active ? '#fff' : '#7a7a90',
    fontSize: 13, fontWeight: 600,
    cursor: 'pointer',
    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 7,
    fontFamily: 'inherit',
    transition: 'all 0.2s',
  }),
  main: {
    flex: 1,
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    paddingTop: 8,
  },
  emptyBox: {
    flex: 1,
    display: 'flex', flexDirection: 'column',
    alignItems: 'center', justifyContent: 'center',
    padding: '24px 20px',
    overflowY: 'auto',
    gap: 18,
    textAlign: 'center',
  },
  bigTitle: {
    fontFamily: "'Syne', sans-serif",
    fontSize: 24, fontWeight: 700,
    color: '#ffffff',
    letterSpacing: '-0.01em',
  },
  bodyText: {
    color: '#8a8aa0',
    fontSize: 14, lineHeight: 1.55,
    maxWidth: 320,
  },
  cta: (disabled) => ({
    width: '100%', maxWidth: 280,
    padding: 13,
    background: 'linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%)',
    border: 'none', borderRadius: 12,
    color: 'white',
    fontFamily: "'Syne', sans-serif",
    fontWeight: 700, fontSize: 15,
    letterSpacing: '0.05em',
    cursor: disabled ? 'not-allowed' : 'pointer',
    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
    boxShadow: disabled ? 'none' : '0 4px 14px rgba(124,109,250,0.3)',
    opacity: disabled ? 0.5 : 1,
    marginTop: 6,
  }),
}

// ═══════════════════════════════════════════════════════════════
// Small components
// ═══════════════════════════════════════════════════════════════

function Logo({ size = 36, fontSize = 11 }) {
  return (
    <div style={{
      width: size, height: size,
      borderRadius: size * 0.3,
      background: 'linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontFamily: "'Syne', sans-serif",
      fontWeight: 800, fontSize, color: 'white', letterSpacing: 0.5,
      boxShadow: '0 4px 14px rgba(124,109,250,0.4), inset 0 1px 0 rgba(255,255,255,0.2)',
      flexShrink: 0,
    }}>AX</div>
  )
}

function Orb({ emoji, color = '#7c6dfa' }) {
  return (
    <div style={{
      position: 'relative',
      width: 80, height: 80,
      display: 'flex', alignItems: 'center', justifyContent: 'center',
    }}>
      <div style={{
        position: 'absolute', inset: -18,
        background: `radial-gradient(circle, ${color}99 0%, transparent 70%)`,
        filter: 'blur(22px)',
        animation: 'axPulse 3s ease-in-out infinite',
      }}/>
      <div style={{
        position: 'relative',
        width: 62, height: 62,
        borderRadius: 18,
        background: 'linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        fontSize: 26,
        boxShadow: '0 10px 36px rgba(124,109,250,0.4), inset 0 1px 0 rgba(255,255,255,0.2)',
        animation: 'axFloat 4s ease-in-out infinite',
      }}>{emoji}</div>
    </div>
  )
}

function TypingDots() {
  return (
    <div style={{ display: 'flex', gap: 5, padding: '2px 0' }}>
      {[0, 1, 2].map(i => (
        <span key={i} style={{
          width: 7, height: 7, borderRadius: '50%',
          background: 'rgba(255,255,255,0.4)',
          display: 'block',
          animation: 'axBounce 1.3s ease infinite',
          animationDelay: `${i * 0.15}s`,
        }}/>
      ))}
    </div>
  )
}

function Message({ msg }) {
  const isUser = msg.role === 'user'
  return (
    <div style={{
      display: 'flex',
      gap: 9,
      alignItems: 'flex-end',
      flexDirection: isUser ? 'row-reverse' : 'row',
      animation: 'axFadeIn 0.3s ease',
    }}>
      {!isUser && <Logo size={28} fontSize={9}/>}
      {isUser && (
        <div style={{
          width: 28, height: 28, borderRadius: '50%',
          background: 'rgba(255,255,255,0.06)',
          border: '1px solid rgba(255,255,255,0.12)',
          color: '#8a8aa0',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontWeight: 700, fontSize: 11, flexShrink: 0,
          fontFamily: "'Syne', sans-serif",
        }}>U</div>
      )}
      <div style={{
        maxWidth: '78%',
        padding: '11px 15px',
        fontSize: 14, lineHeight: 1.6,
        borderRadius: 18,
        borderBottomRightRadius: isUser ? 5 : 18,
        borderBottomLeftRadius: isUser ? 18 : 5,
        background: isUser
          ? 'linear-gradient(135deg, #7c6dfa 0%, #6c5ce7 100%)'
          : (msg.isError ? 'rgba(250,109,154,0.1)' : 'rgba(255,255,255,0.04)'),
        border: isUser
          ? 'none'
          : `1px solid ${msg.isError ? 'rgba(250,109,154,0.3)' : 'rgba(255,255,255,0.08)'}`,
        color: isUser ? 'white' : (msg.isError ? '#fa6d9a' : '#e4e4f0'),
        boxShadow: isUser ? '0 4px 16px rgba(124,109,250,0.25)' : 'none',
      }}>
        {msg.loading
          ? <TypingDots/>
          : <div className="ax-content" dangerouslySetInnerHTML={{ __html: parseMarkdown(msg.content) }}/>
        }
        {msg.timestamp && !msg.loading && (
          <div style={{
            fontSize: 10, marginTop: 5, opacity: 0.4,
            textAlign: isUser ? 'right' : 'left',
          }}>{msg.timestamp}</div>
        )}
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════
// Settings screen
// ═══════════════════════════════════════════════════════════════

function Settings({ onSave, onClose, hasKey }) {
  const [key, setKey] = useState(() => localStorage.getItem('axiom_api_key') || '')
  const [visible, setVisible] = useState(false)

  const save = () => {
    const k = key.trim()
    if (!k) return
    localStorage.setItem('axiom_api_key', k)
    onSave(k)
  }

  return (
    <div style={{
      position: 'fixed', inset: 0,
      background: 'rgba(6,6,8,0.85)',
      backdropFilter: 'blur(20px)',
      WebkitBackdropFilter: 'blur(20px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      padding: 20, zIndex: 100,
    }}>
      <div style={{
        width: '100%', maxWidth: 370,
        background: 'rgba(20,20,30,0.9)',
        border: '1px solid rgba(255,255,255,0.1)',
        borderRadius: 22,
        padding: '34px 24px 24px',
        position: 'relative',
        overflow: 'hidden',
        boxShadow: '0 20px 60px rgba(0,0,0,0.5)',
      }}>
        <div style={{
          position: 'absolute',
          top: -100, left: '50%',
          transform: 'translateX(-50%)',
          width: 280, height: 280,
          background: 'radial-gradient(circle, rgba(124,109,250,0.35) 0%, transparent 70%)',
          filter: 'blur(40px)',
          pointerEvents: 'none',
        }}/>

        <div style={{ textAlign: 'center', marginBottom: 26, position: 'relative' }}>
          <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 12 }}>
            <Logo size={56} fontSize={16}/>
          </div>
          <div style={{
            fontFamily: "'Syne', sans-serif",
            fontWeight: 700, fontSize: 24,
            letterSpacing: '0.04em',
            color: '#ffffff',
          }}>AXIOM</div>
          <div style={{
            fontSize: 10, color: '#5a5a70',
            textTransform: 'uppercase', letterSpacing: '0.15em',
            marginTop: 4,
          }}>AI Agent · Powered by Claude</div>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 12, position: 'relative' }}>
          <div style={{ fontSize: 12, color: '#8a8aa0', fontWeight: 500 }}>
            Anthropic API Key
          </div>

          <div style={{ display: 'flex', gap: 8 }}>
            <input
              type={visible ? 'text' : 'password'}
              placeholder="sk-ant-..."
              value={key}
              onChange={e => setKey(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && save()}
              autoComplete="off"
              spellCheck={false}
              style={{
                flex: 1,
                background: 'rgba(0,0,0,0.3)',
                border: '1px solid rgba(255,255,255,0.1)',
                borderRadius: 11,
                padding: '11px 13px',
                color: '#e8e8f0',
                fontFamily: 'inherit', fontSize: 14,
                outline: 'none',
              }}
            />
            <button
              onClick={() => setVisible(v => !v)}
              style={{
                background: 'rgba(0,0,0,0.3)',
                border: '1px solid rgba(255,255,255,0.1)',
                borderRadius: 11,
                width: 44, height: 44,
                cursor: 'pointer', fontSize: 16,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                flexShrink: 0,
              }}
            >{visible ? '🙈' : '👁️'}</button>
          </div>

          <button onClick={save} disabled={!key.trim()} style={S.cta(!key.trim())}>
            {hasKey ? 'Save Changes' : 'Launch AXIOM'}
            <span>→</span>
          </button>

          {hasKey && (
            <button
              onClick={onClose}
              style={{
                width: '100%', padding: 11,
                background: 'transparent',
                border: '1px solid rgba(255,255,255,0.1)',
                borderRadius: 11,
                color: '#8a8aa0',
                fontFamily: 'inherit', fontWeight: 500, fontSize: 13,
                cursor: 'pointer',
              }}
            >Cancel</button>
          )}

          <a
            href="https://console.anthropic.com/settings/keys"
            target="_blank" rel="noreferrer"
            style={{
              color: '#6a6a80', fontSize: 12,
              textAlign: 'center', textDecoration: 'none',
              marginTop: 2,
            }}
          >Get your API key →</a>
        </div>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════
// Chat tab
// ═══════════════════════════════════════════════════════════════

function ChatTab({ apiKey }) {
  const [messages, setMessages] = useState([])
  const [history, setHistory] = useState([])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const bottomRef = useRef(null)
  const taRef = useRef(null)

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  const updateLast = useCallback((fn) => {
    setMessages(prev => {
      if (!prev.length) return prev
      const copy = [...prev]
      copy[copy.length - 1] = fn(copy[copy.length - 1])
      return copy
    })
  }, [])

  const send = async (text) => {
    const userText = (text ?? input).trim()
    if (!userText || loading) return
    setInput('')
    if (taRef.current) taRef.current.style.height = '24px'

    const now = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })
    setMessages(prev => [
      ...prev,
      { role: 'user', content: userText, timestamp: now },
      { role: 'assistant', content: '', loading: true },
    ])
    const newHistory = [...history, { role: 'user', content: userText }]
    setHistory(newHistory)
    setLoading(true)

    try {
      const reply = await chat(apiKey, newHistory)
      const ts = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })
      updateLast(() => ({ role: 'assistant', content: reply, timestamp: ts, loading: false }))
      setHistory([...newHistory, { role: 'assistant', content: reply }])
    } catch (err) {
      updateLast(() => ({ role: 'assistant', content: `⚠️ ${err.message}`, loading: false, isError: true }))
    } finally {
      setLoading(false)
    }
  }

  const handleKey = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send() }
  }

  const clear = () => { setMessages([]); setHistory([]) }
  const isEmpty = messages.length === 0

  return (
    <>
      {isEmpty ? (
        <div style={S.emptyBox}>
          <Orb emoji="⚡"/>
          <div style={S.bigTitle}>What can I do for you?</div>
          <div style={S.bodyText}>
            Daily tasks, research, coding, analysis — ask me anything.
          </div>
          <div style={{
            display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 8,
            width: '100%', maxWidth: 440, marginTop: 8,
          }}>
            {SUGGESTED.map((s, i) => (
              <button key={i} onClick={() => send(s.text)} style={{
                background: 'rgba(255,255,255,0.03)',
                border: '1px solid rgba(255,255,255,0.06)',
                borderRadius: 11,
                padding: '11px 12px',
                textAlign: 'left', cursor: 'pointer',
                color: '#d8d8e8',
                fontFamily: 'inherit', fontSize: 12,
                display: 'flex', alignItems: 'flex-start', gap: 8,
                lineHeight: 1.35,
              }}>
                <span style={{ fontSize: 14, flexShrink: 0, marginTop: 1 }}>{s.icon}</span>
                <span>{s.text}</span>
              </button>
            ))}
          </div>
        </div>
      ) : (
        <div style={{
          flex: 1, overflowY: 'auto',
          padding: '4px 12px 8px',
          display: 'flex', flexDirection: 'column', gap: 12,
          WebkitOverflowScrolling: 'touch',
        }}>
          {messages.map((msg, i) => <Message key={i} msg={msg}/>)}
          <div ref={bottomRef}/>
        </div>
      )}

      <div style={{
        padding: '10px 12px',
        paddingBottom: 'calc(10px + env(safe-area-inset-bottom, 0px))',
        flexShrink: 0,
        position: 'relative',
      }}>
        {!isEmpty && (
          <button onClick={clear} style={{
            position: 'absolute', top: -26, left: '50%',
            transform: 'translateX(-50%)',
            background: 'rgba(255,255,255,0.04)',
            border: '1px solid rgba(255,255,255,0.08)',
            color: '#6a6a80',
            borderRadius: 100,
            padding: '4px 11px',
            fontSize: 11, cursor: 'pointer',
            display: 'flex', alignItems: 'center', gap: 5,
          }}>✕ Clear</button>
        )}
        <div style={{
          display: 'flex', alignItems: 'flex-end', gap: 8,
          background: 'rgba(255,255,255,0.04)',
          border: '1px solid rgba(255,255,255,0.08)',
          borderRadius: 16,
          padding: '8px 8px 8px 14px',
        }}>
          <textarea
            ref={taRef}
            rows={1}
            value={input}
            onChange={e => {
              setInput(e.target.value)
              e.target.style.height = 'auto'
              e.target.style.height = Math.min(e.target.scrollHeight, 120) + 'px'
            }}
            onKeyDown={handleKey}
            placeholder="Ask AXIOM anything…"
            disabled={loading}
            style={{
              flex: 1,
              background: 'transparent', border: 'none', outline: 'none',
              color: '#e8e8f0',
              fontFamily: 'inherit', fontSize: 15,
              resize: 'none',
              minHeight: 24, maxHeight: 120,
              lineHeight: 1.5,
              overflowY: 'auto',
              padding: '5px 0',
            }}
          />
          <button
            onClick={() => send()}
            disabled={loading || !input.trim()}
            style={{
              width: 34, height: 34,
              borderRadius: 10, border: 'none',
              background: 'linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%)',
              color: 'white', cursor: 'pointer',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
              fontSize: 16, fontWeight: 'bold',
              opacity: (loading || !input.trim()) ? 0.3 : 1,
              boxShadow: '0 2px 8px rgba(124,109,250,0.3)',
            }}
          >↑</button>
        </div>
      </div>
    </>
  )
}

// ═══════════════════════════════════════════════════════════════
// Email tab
// ═══════════════════════════════════════════════════════════════

function EmailTab({ apiKey }) {
  const [account, setAccount] = useState(null)
  const [emails, setEmails] = useState([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)
  const [drafting, setDrafting] = useState(null)
  const [connecting, setConnecting] = useState(false)

  useEffect(() => {
    (async () => {
      try {
        await initMsal()
        const acc = await getAccount()
        if (acc) {
          setAccount(acc)
          loadEmails()
        }
      } catch (err) {
        setError(err.message)
      }
    })()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const loadEmails = async () => {
    setLoading(true)
    setError(null)
    try {
      const list = await listEmails(15)
      setEmails(list)
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  const handleConnect = async () => {
    setConnecting(true)
    setError(null)
    try {
      const acc = await signIn()
      if (acc) {
        setAccount(acc)
        await loadEmails()
      }
    } catch (err) {
      setError(err.message)
    } finally {
      setConnecting(false)
    }
  }

  const handleDisconnect = async () => {
    try {
      await signOut()
      setAccount(null)
      setEmails([])
    } catch (err) {
      setError(err.message)
    }
  }

  const handleDraft = async (email) => {
    setDrafting(email.id)
    setError(null)
    try {
      const full = await getEmailBody(email.id)
      const bodyText = stripHtml(full.body?.content || full.bodyPreview || '')

      const replyText = await draftEmailReply(apiKey, {
        from: full.from?.emailAddress?.name || full.from?.emailAddress?.address || 'Unknown',
        subject: full.subject || '(no subject)',
        body: bodyText,
      })

      await createDraftReply(email.id, replyText)
      setEmails(prev => prev.map(e => e.id === email.id ? { ...e, _drafted: true } : e))
    } catch (err) {
      setError(err.message)
    } finally {
      setDrafting(null)
    }
  }

  if (!account) {
    return (
      <div style={S.emptyBox}>
        <Orb emoji="📧" color="#00b7c3"/>
        <div style={S.bigTitle}>Connect Your Outlook</div>
        <div style={S.bodyText}>
          AXIOM reads your inbox, drafts replies in your voice, and saves them to your Outlook drafts — you review &amp; send.
        </div>
        <button onClick={handleConnect} disabled={connecting} style={S.cta(connecting)}>
          {connecting ? 'Connecting…' : 'Connect Outlook'}
          <span>→</span>
        </button>
        {error && (
          <div style={{
            background: 'rgba(250,109,154,0.08)',
            border: '1px solid rgba(250,109,154,0.3)',
            color: '#fa6d9a',
            borderRadius: 10,
            padding: '10px 14px',
            fontSize: 13, lineHeight: 1.5,
            maxWidth: '100%', wordBreak: 'break-word',
          }}>⚠️ {error}</div>
        )}
        <div style={{
          fontSize: 11, color: '#5a5a70',
          maxWidth: 280, lineHeight: 1.5,
        }}>
          Your emails stay between you and Microsoft.
          AXIOM only accesses your inbox while you're using the app.
        </div>
      </div>
    )
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        padding: '10px 16px 11px',
        margin: '0 14px',
        borderBottom: '1px solid rgba(255,255,255,0.05)',
        flexShrink: 0,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 }}>
          <div style={{
            width: 34, height: 34, borderRadius: '50%',
            background: 'linear-gradient(135deg, #00b7c3 0%, #7c6dfa 100%)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            color: 'white', fontWeight: 700, fontSize: 14,
            flexShrink: 0,
          }}>{(account.username || '?')[0].toUpperCase()}</div>
          <div style={{ minWidth: 0 }}>
            <div style={{
              fontSize: 13, fontWeight: 600, color: '#e8e8f0',
              whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
              maxWidth: 180,
            }}>{account.name || 'Account'}</div>
            <div style={{
              fontSize: 11, color: '#6a6a80',
              whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
              maxWidth: 180,
            }}>{account.username}</div>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 6 }}>
          <button
            onClick={loadEmails}
            disabled={loading}
            style={{ ...S.iconBtn, width: 32, height: 32, fontSize: 14 }}
          >{loading ? '…' : '↻'}</button>
          <button
            onClick={handleDisconnect}
            style={{ ...S.iconBtn, width: 32, height: 32, fontSize: 14 }}
          >⎋</button>
        </div>
      </div>

      {error && (
        <div style={{
          margin: '10px 14px',
          background: 'rgba(250,109,154,0.08)',
          border: '1px solid rgba(250,109,154,0.3)',
          color: '#fa6d9a',
          borderRadius: 10,
          padding: '10px 14px',
          fontSize: 13, lineHeight: 1.5,
          wordBreak: 'break-word',
        }}>⚠️ {error}</div>
      )}

      {loading && emails.length === 0 ? (
        <div style={{
          flex: 1,
          display: 'flex', flexDirection: 'column',
          alignItems: 'center', justifyContent: 'center', gap: 14,
        }}>
          <TypingDots/>
          <div style={{ color: '#6a6a80', fontSize: 13 }}>Loading your inbox…</div>
        </div>
      ) : emails.length === 0 ? (
        <div style={{
          flex: 1,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          color: '#5a5a70', fontSize: 14,
        }}>No emails found.</div>
      ) : (
        <div style={{
          flex: 1, overflowY: 'auto',
          padding: '12px 14px calc(14px + env(safe-area-inset-bottom, 0px))',
          display: 'flex', flexDirection: 'column', gap: 9,
          WebkitOverflowScrolling: 'touch',
        }}>
          {emails.map(email => (
            <div key={email.id} style={{
              background: !email.isRead ? 'rgba(124,109,250,0.06)' : 'rgba(255,255,255,0.03)',
              border: `1px solid ${!email.isRead ? 'rgba(124,109,250,0.2)' : 'rgba(255,255,255,0.06)'}`,
              borderRadius: 13,
              padding: '12px 14px',
              position: 'relative',
            }}>
              {!email.isRead && (
                <div style={{
                  position: 'absolute', top: 16, left: 5,
                  width: 4, height: 4, borderRadius: '50%',
                  background: '#7c6dfa',
                  boxShadow: '0 0 8px #7c6dfa',
                }}/>
              )}
              <div style={{
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                gap: 10, marginBottom: 4,
              }}>
                <div style={{
                  fontSize: 13, fontWeight: 600, color: '#e8e8f0',
                  whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                  flex: 1,
                }}>
                  {email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown'}
                </div>
                <div style={{ fontSize: 11, color: '#5a5a70', flexShrink: 0 }}>
                  {formatRelative(email.receivedDateTime)}
                </div>
              </div>
              <div style={{
                fontSize: 13.5, color: '#c8c8dc',
                marginBottom: 5, lineHeight: 1.4,
                wordBreak: 'break-word',
              }}>{email.subject || '(no subject)'}</div>
              <div style={{
                fontSize: 12, color: '#6a6a80',
                lineHeight: 1.5, marginBottom: 10,
                display: '-webkit-box',
                WebkitLineClamp: 2,
                WebkitBoxOrient: 'vertical',
                overflow: 'hidden',
              }}>{email.bodyPreview?.slice(0, 140)}…</div>
              <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                {email._drafted ? (
                  <div style={{
                    background: 'rgba(109,250,204,0.1)',
                    border: '1px solid rgba(109,250,204,0.3)',
                    color: '#6dfacc',
                    borderRadius: 9,
                    padding: '7px 11px',
                    fontSize: 12, fontWeight: 500,
                  }}>✓ Draft saved — check Outlook</div>
                ) : (
                  <button
                    onClick={() => handleDraft(email)}
                    disabled={drafting === email.id}
                    style={{
                      background: 'linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%)',
                      border: 'none', borderRadius: 9,
                      padding: '7px 13px',
                      color: 'white',
                      fontFamily: 'inherit',
                      fontSize: 12.5, fontWeight: 600,
                      cursor: 'pointer',
                      display: 'flex', alignItems: 'center', gap: 6,
                      boxShadow: '0 2px 10px rgba(124,109,250,0.3)',
                      opacity: drafting === email.id ? 0.7 : 1,
                    }}
                  >
                    {drafting === email.id ? (
                      <>
                        <span style={{
                          width: 12, height: 12,
                          border: '2px solid rgba(255,255,255,0.3)',
                          borderTopColor: 'white',
                          borderRadius: '50%',
                          animation: 'axSpin 0.7s linear infinite',
                          display: 'inline-block',
                        }}/>
                        <span>Drafting…</span>
                      </>
                    ) : (
                      <><span>✨</span><span>Generate Draft</span></>
                    )}
                  </button>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════
// Main App
// ═══════════════════════════════════════════════════════════════

export default function App() {
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('axiom_api_key') || '')
  const [showSettings, setShow] = useState(false)
  const [tab, setTab] = useState('chat')

  if (!apiKey) {
    return (
      <>
        <GlobalStyles/>
        <Settings onSave={k => { setApiKey(k); setShow(false) }} hasKey={false}/>
      </>
    )
  }

  return (
    <>
      <GlobalStyles/>
      <BackgroundLayers/>

      <div style={S.app}>
        {/* Header */}
        <header style={S.header}>
          <div style={S.headerLeft}>
            <Logo size={36} fontSize={11}/>
            <div>
              <div style={S.brand}>AXIOM</div>
              <div style={S.brandSub}>AI Agent</div>
            </div>
          </div>
          <button onClick={() => setShow(true)} style={S.iconBtn}>⚙</button>
        </header>

        {/* Tabs */}
        <div style={S.tabs}>
          <button onClick={() => setTab('chat')} style={S.tabBtn(tab === 'chat')}>
            <span style={{ fontSize: 14 }}>💬</span>
            <span>Chat</span>
          </button>
          <button onClick={() => setTab('email')} style={S.tabBtn(tab === 'email')}>
            <span style={{ fontSize: 14 }}>📧</span>
            <span>Email</span>
          </button>
        </div>

        {/* Body */}
        <main style={S.main}>
          {tab === 'chat' ? <ChatTab apiKey={apiKey}/> : <EmailTab apiKey={apiKey}/>}
        </main>
      </div>

      {showSettings && (
        <Settings
          onSave={k => { setApiKey(k); setShow(false) }}
          onClose={() => setShow(false)}
          hasKey={true}
        />
      )}
    </>
  )
}

// ═══════════════════════════════════════════════════════════════
// Background layers
// ═══════════════════════════════════════════════════════════════

function BackgroundLayers() {
  return (
    <>
      <div style={{
        position: 'fixed', inset: 0, zIndex: 0,
        background: `
          radial-gradient(ellipse 80% 60% at 50% 0%, #1a0b2e 0%, transparent 50%),
          radial-gradient(ellipse 80% 60% at 50% 100%, #0a0a1a 0%, transparent 50%),
          #060608
        `,
      }}/>
      <div style={{
        position: 'fixed', inset: 0, zIndex: 0,
        backgroundImage: 'linear-gradient(rgba(255,255,255,0.025) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.025) 1px, transparent 1px)',
        backgroundSize: '48px 48px',
        pointerEvents: 'none',
      }}/>
      <div style={{
        position: 'fixed', borderRadius: '50%',
        width: 460, height: 460,
        top: -180, left: -180,
        filter: 'blur(90px)',
        pointerEvents: 'none', zIndex: 0,
        opacity: 0.3,
        background: 'radial-gradient(circle, #7c6dfa 0%, transparent 70%)',
        animation: 'axDrift1 20s ease-in-out infinite',
      }}/>
      <div style={{
        position: 'fixed', borderRadius: '50%',
        width: 460, height: 460,
        bottom: -180, right: -180,
        filter: 'blur(90px)',
        pointerEvents: 'none', zIndex: 0,
        opacity: 0.3,
        background: 'radial-gradient(circle, #fa6d9a 0%, transparent 70%)',
        animation: 'axDrift2 24s ease-in-out infinite',
      }}/>
    </>
  )
}

// ═══════════════════════════════════════════════════════════════
// Global styles (fonts, animations, markdown)
// ═══════════════════════════════════════════════════════════════

function GlobalStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=Inter:wght@300;400;500;600;700&display=swap');

      *, *::before, *::after {
        box-sizing: border-box;
        margin: 0; padding: 0;
        -webkit-tap-highlight-color: transparent;
      }

      html, body {
        height: 100%;
        overflow: hidden;
        background: #060608;
        color: #e8e8f0;
        font-family: 'Inter', -apple-system, sans-serif;
        -webkit-font-smoothing: antialiased;
      }

      #root {
        height: 100%;
        display: flex;
        flex-direction: column;
      }

      button { font-family: inherit; }
      input::placeholder, textarea::placeholder { color: #4a4a5c; }

      @keyframes axPulse {
        0%,100% { opacity: 0.6; transform: scale(1); }
        50% { opacity: 1; transform: scale(1.1); }
      }

      @keyframes axFloat {
        0%,100% { transform: translateY(0); }
        50% { transform: translateY(-7px); }
      }

      @keyframes axBounce {
        0%,60%,100% { transform: translateY(0); opacity: 0.4; }
        30% { transform: translateY(-6px); opacity: 1; }
      }

      @keyframes axFadeIn {
        from { opacity: 0; transform: translateY(8px); }
        to { opacity: 1; transform: translateY(0); }
      }

      @keyframes axDrift1 {
        0%,100% { transform: translate(0,0) scale(1); }
        50% { transform: translate(60px, 40px) scale(1.12); }
      }

      @keyframes axDrift2 {
        0%,100% { transform: translate(0,0) scale(1); }
        50% { transform: translate(-60px, -40px) scale(1.15); }
      }

      @keyframes axSpin { to { transform: rotate(360deg); } }

      .ax-content strong { color: #fff; font-weight: 600; }
      .ax-content em { font-style: italic; color: #c0c0dc; }
      .ax-h1 { font-family: 'Syne',sans-serif; font-size: 17px; font-weight: 700; margin: 12px 0 5px; color: #fff; }
      .ax-h2 { font-family: 'Syne',sans-serif; font-size: 15px; font-weight: 600; margin: 10px 0 4px; color: #fff; }
      .ax-h3 { font-size: 14px; font-weight: 600; margin: 8px 0 3px; color: #fff; }
      .ax-li { padding: 3px 0 3px 16px; position: relative; list-style: none; display: block; }
      .ax-li::before { content: '›'; position: absolute; left: 0; color: #7c6dfa; font-weight: 700; }
      .ax-li-n { padding: 3px 0 3px 4px; display: block; }
      .ax-gap { height: 8px; }

      .ax-code {
        background: rgba(0,0,0,0.4);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 10px;
        padding: 12px;
        overflow-x: auto;
        font-family: 'SF Mono','Courier New',monospace;
        font-size: 12px; margin: 8px 0;
        color: #6dfacc; line-height: 1.5;
        white-space: pre;
      }

      .ax-inline {
        background: rgba(0,0,0,0.35);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 5px;
        padding: 1px 6px;
        font-family: 'SF Mono',monospace;
        font-size: 12.5px; color: #fa6d9a;
      }

      ::-webkit-scrollbar { width: 3px; }
      ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); border-radius: 3px; }
      ::-webkit-scrollbar-track { background: transparent; }
    `}</style>
  )
}
