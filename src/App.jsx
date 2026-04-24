import { useState, useRef, useEffect, useCallback } from 'react'
import { chat, draftEmailReply } from './claude.js'
import {
  initMsal, signIn, signOut, getAccount,
  listEmails, getEmailBody, createDraftReply, stripHtml,
} from './outlook.js'

// ═══════════════════════════════════════════════════════════════════
// Utilities
// ═══════════════════════════════════════════════════════════════════

function parseMarkdown(text) {
  if (!text) return ''
  return text
    .replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) =>
      `<pre class="ax-code"><code>${code.trim().replace(/</g,'&lt;').replace(/>/g,'&gt;')}</code></pre>`
    )
    .replace(/`([^`\n]+)`/g, '<code class="ax-inline">$1</code>')
    .replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\*([^*\n]+)\*/g, '<em>$1</em>')
    .replace(/^### (.+)$/gm, '<h3 class="ax-h3">$1</h3>')
    .replace(/^## (.+)$/gm,  '<h2 class="ax-h2">$1</h2>')
    .replace(/^# (.+)$/gm,   '<h1 class="ax-h1">$1</h1>')
    .replace(/^[-•] (.+)$/gm,'<li class="ax-li">$1</li>')
    .replace(/^\d+\. (.+)$/gm,'<li class="ax-li-n">$1</li>')
    .replace(/\n\n/g, '<div class="ax-gap"></div>')
    .replace(/\n/g, '<br/>')
}

function formatRelative(iso) {
  const d = new Date(iso)
  const diff = (Date.now() - d.getTime()) / 1000
  if (diff < 60) return 'just now'
  if (diff < 3600) return `${Math.floor(diff/60)}m ago`
  if (diff < 86400) return `${Math.floor(diff/3600)}h ago`
  if (diff < 604800) return `${Math.floor(diff/86400)}d ago`
  return d.toLocaleDateString()
}

const SUGGESTED = [
  { icon: '📈', text: 'Analyze a TASI stock for me' },
  { icon: '✍️', text: 'Draft a professional email' },
  { icon: '💻', text: 'Help me write a script' },
  { icon: '🧠', text: 'Explain something complex' },
  { icon: '📊', text: 'Pros and cons of my idea' },
  { icon: '🗓️', text: 'Help me plan my week' },
]

// ═══════════════════════════════════════════════════════════════════
// Mini components
// ═══════════════════════════════════════════════════════════════════

function TypingDots() {
  return (
    <div className="typing">
      <span/><span/><span/>
    </div>
  )
}

function Logo({ size = 'md' }) {
  const sz = size === 'lg' ? 64 : size === 'sm' ? 28 : 38
  const fs = size === 'lg' ? 18 : size === 'sm' ? 9 : 11
  return (
    <div className="logo-mark" style={{ width: sz, height: sz, fontSize: fs }}>
      <span>AX</span>
    </div>
  )
}

function Message({ msg }) {
  const isUser = msg.role === 'user'
  return (
    <div className={`msg-row ${isUser ? 'user' : 'agent'}`}>
      {!isUser && <Logo size="sm"/>}
      <div className={`bubble ${isUser ? 'user' : 'agent'} ${msg.isError ? 'error' : ''}`}>
        {msg.loading
          ? <TypingDots/>
          : <div className="msg-content" dangerouslySetInnerHTML={{ __html: parseMarkdown(msg.content) }}/>
        }
        {msg.timestamp && !msg.loading && (
          <div className="ts">{msg.timestamp}</div>
        )}
      </div>
      {isUser && <div className="user-avatar">U</div>}
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════
// Settings / API key screen
// ═══════════════════════════════════════════════════════════════════

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
    <div className="settings-overlay">
      <div className="settings-card">
        <div className="settings-glow"/>

        <div className="settings-header">
          <Logo size="lg"/>
          <div className="settings-brand">AXIOM</div>
          <div className="settings-sub">AI Agent · Powered by Claude</div>
        </div>

        <div className="settings-form">
          <div className="field-label">Anthropic API Key</div>
          <div className="key-row">
            <input
              type={visible ? 'text' : 'password'}
              placeholder="sk-ant-..."
              value={key}
              onChange={e => setKey(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && save()}
              autoComplete="off"
              spellCheck={false}
            />
            <button className="eye-btn" onClick={() => setVisible(v => !v)}>
              {visible ? '🙈' : '👁️'}
            </button>
          </div>

          <button className="cta" onClick={save} disabled={!key.trim()}>
            {hasKey ? 'Save Changes' : 'Launch AXIOM'}
            <span className="cta-arrow">→</span>
          </button>

          {hasKey && (
            <button className="cta-secondary" onClick={onClose}>Cancel</button>
          )}

          <a
            href="https://console.anthropic.com/settings/keys"
            target="_blank" rel="noreferrer"
            className="get-key-link"
          >
            Get your API key →
          </a>
        </div>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════
// Chat tab
// ═══════════════════════════════════════════════════════════════════

function ChatTab({ apiKey }) {
  const [messages, setMessages] = useState([])
  const [history, setHistory]   = useState([])
  const [input, setInput]       = useState('')
  const [loading, setLoading]   = useState(false)
  const bottomRef = useRef(null)
  const taRef     = useRef(null)

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

    const now = new Date().toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit' })
    setMessages(prev => [...prev, { role:'user', content: userText, timestamp: now }, { role:'assistant', content: '', loading: true }])
    const newHistory = [...history, { role:'user', content: userText }]
    setHistory(newHistory)
    setLoading(true)

    try {
      const reply = await chat(apiKey, newHistory)
      const ts = new Date().toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit' })
      updateLast(() => ({ role:'assistant', content: reply, timestamp: ts, loading: false }))
      setHistory([...newHistory, { role:'assistant', content: reply }])
    } catch (err) {
      updateLast(() => ({ role:'assistant', content:`⚠️ ${err.message}`, loading: false, isError: true }))
    } finally {
      setLoading(false)
    }
  }

  const handleKey = (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send() } }
  const clear = () => { setMessages([]); setHistory([]) }
  const isEmpty = messages.length === 0

  return (
    <>
      {isEmpty ? (
        <div className="empty">
          <div className="empty-orb">
            <div className="orb-glow"/>
            <div className="orb-inner">⚡</div>
          </div>
          <div style={{
            fontFamily: "'Syne', sans-serif",
            fontSize: 26, fontWeight: 700, color: '#ffffff',
            letterSpacing: '-0.01em', textAlign: 'center',
          }}>What can I do for you?</div>
          <div style={{
            color: '#8a8aa0', fontSize: 14, lineHeight: 1.6,
            maxWidth: 320, textAlign: 'center', marginTop: -8,
          }}>Daily tasks, research, coding, analysis — ask me anything.</div>
          <div className="suggestions">
            {SUGGESTED.map((s,i) => (
              <button key={i} className="sug-btn" onClick={() => send(s.text)}>
                <span className="sug-icon">{s.icon}</span>
                <span>{s.text}</span>
              </button>
            ))}
          </div>
        </div>
      ) : (
        <div className="messages">
          {messages.map((msg,i) => <Message key={i} msg={msg}/>)}
          <div ref={bottomRef}/>
        </div>
      )}

      <div className="composer">
        {!isEmpty && (
          <button className="clear-chat-btn" onClick={clear}>
            <span>✕</span> Clear
          </button>
        )}
        <div className="input-wrap">
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
          />
          <button
            className="send-btn"
            onClick={() => send()}
            disabled={loading || !input.trim()}
            aria-label="Send"
          >↑</button>
        </div>
      </div>
    </>
  )
}

// ═══════════════════════════════════════════════════════════════════
// Email tab
// ═══════════════════════════════════════════════════════════════════

function EmailTab({ apiKey }) {
  const [account, setAccount]     = useState(null)
  const [emails, setEmails]       = useState([])
  const [loading, setLoading]     = useState(false)
  const [error, setError]         = useState(null)
  const [drafting, setDrafting]   = useState(null)
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
      <div style={{
        flex: 1,
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '28px 24px',
        overflowY: 'auto',
        gap: 18,
        textAlign: 'center',
      }}>
        <div className="empty-orb">
          <div className="orb-glow" style={{background: 'radial-gradient(circle, rgba(0,183,195,0.6) 0%, transparent 70%)'}}/>
          <div className="orb-inner">📧</div>
        </div>

        <div style={{
          fontFamily: "'Syne', sans-serif",
          fontSize: 26,
          fontWeight: 700,
          color: '#ffffff',
          letterSpacing: '-0.01em',
          marginTop: 4,
        }}>
          Connect Your Outlook
        </div>

        <div style={{
          color: '#8a8aa0',
          fontSize: 14,
          lineHeight: 1.6,
          maxWidth: 320,
          marginTop: -6,
        }}>
          AXIOM reads your inbox, drafts replies in your voice, and saves them to your Outlook drafts — you review &amp; send.
        </div>

        <button
          onClick={handleConnect}
          disabled={connecting}
          style={{
            width: '100%',
            maxWidth: 280,
            padding: 14,
            background: 'linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%)',
            border: 'none',
            borderRadius: 12,
            color: 'white',
            fontFamily: "'Syne', sans-serif",
            fontWeight: 700,
            fontSize: 15,
            letterSpacing: '0.05em',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: 8,
            boxShadow: '0 4px 14px rgba(124,109,250,0.3)',
            opacity: connecting ? 0.6 : 1,
            marginTop: 8,
          }}
        >
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
            fontSize: 13,
            lineHeight: 1.5,
            maxWidth: '100%',
            wordBreak: 'break-word',
          }}>
            ⚠️ {error}
          </div>
        )}

        <div style={{
          fontSize: 11,
          color: '#5a5a70',
          maxWidth: 280,
          lineHeight: 1.5,
          marginTop: 8,
        }}>
          Your emails stay between you and Microsoft.
          AXIOM only accesses your inbox when you're using the app.
        </div>
      </div>
    )
  }

  return (
    <div className="email-main">
      <div className="account-bar">
        <div className="account-info">
          <div className="account-avatar">{(account.username || '?')[0].toUpperCase()}</div>
          <div>
            <div className="account-name">{account.name || 'Account'}</div>
            <div className="account-email">{account.username}</div>
          </div>
        </div>
        <div className="account-actions">
          <button className="icon-ghost" onClick={loadEmails} disabled={loading} title="Refresh">
            {loading ? '…' : '↻'}
          </button>
          <button className="icon-ghost" onClick={handleDisconnect} title="Disconnect">⎋</button>
        </div>
      </div>

      {error && <div className="error-box" style={{margin:'10px 14px'}}>⚠️ {error}</div>}

      {loading && emails.length === 0 ? (
        <div className="email-loading">
          <TypingDots/>
          <div className="loading-text">Loading your inbox…</div>
        </div>
      ) : emails.length === 0 ? (
        <div className="empty-inbox">No emails found.</div>
      ) : (
        <div className="email-list">
          {emails.map(email => (
            <div key={email.id} className={`email-card ${!email.isRead ? 'unread' : ''}`}>
              <div className="email-head">
                <div className="email-from">
                  {email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown'}
                </div>
                <div className="email-time">{formatRelative(email.receivedDateTime)}</div>
              </div>
              <div className="email-subject">{email.subject || '(no subject)'}</div>
              <div className="email-preview">{email.bodyPreview?.slice(0, 140)}…</div>
              <div className="email-actions">
                {email._drafted ? (
                  <div className="drafted-pill">
                    ✓ Draft saved — check Outlook
                  </div>
                ) : (
                  <button
                    className="draft-btn"
                    onClick={() => handleDraft(email)}
                    disabled={drafting === email.id}
                  >
                    {drafting === email.id ? (
                      <>
                        <span className="spinner"/>
                        <span>Drafting…</span>
                      </>
                    ) : (
                      <>
                        <span>✨</span>
                        <span>Generate Draft</span>
                      </>
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

// ═══════════════════════════════════════════════════════════════════
// Main App
// ═══════════════════════════════════════════════════════════════════

export default function App() {
  const [apiKey, setApiKey]       = useState(() => localStorage.getItem('axiom_api_key') || '')
  const [showSettings, setShow]   = useState(false)
  const [tab, setTab]             = useState('chat')

  if (!apiKey) {
    return <Settings onSave={k => { setApiKey(k); setShow(false) }} hasKey={false}/>
  }

  return (
    <div className="app">
      <div className="bg-gradient"/>
      <div className="bg-grid"/>
      <div className="bg-glow bg-glow-1"/>
      <div className="bg-glow bg-glow-2"/>

      <header className="header">
        <div className="header-left">
          <Logo/>
          <div>
            <div className="brand">AXIOM</div>
            <div className="brand-sub">AI Agent</div>
          </div>
        </div>

        <div className="header-right">
          <button className="icon-ghost" onClick={() => setShow(true)} title="Settings">⚙</button>
        </div>
      </header>

      <div className="tabs">
        <button className={`tab ${tab === 'chat' ? 'active' : ''}`} onClick={() => setTab('chat')}>
          <span className="tab-icon">💬</span>
          <span>Chat</span>
        </button>
        <button className={`tab ${tab === 'email' ? 'active' : ''}`} onClick={() => setTab('email')}>
          <span className="tab-icon">📧</span>
          <span>Email</span>
        </button>
      </div>

      <main className="main">
        {tab === 'chat' ? <ChatTab apiKey={apiKey}/> : <EmailTab apiKey={apiKey}/>}
      </main>

      {showSettings && (
        <Settings
          onSave={k => { setApiKey(k); setShow(false) }}
          onClose={() => setShow(false)}
          hasKey={true}
        />
      )}

      <style>{styles}</style>
    </div>
  )
}

// ═══════════════════════════════════════════════════════════════════
// Styles
// ═══════════════════════════════════════════════════════════════════

const styles = `
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=Inter:wght@300;400;500;600;700&display=swap');

*, *::before, *::after {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
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

.app {
  display: flex;
  flex-direction: column;
  height: 100%;
  max-width: 720px;
  margin: 0 auto;
  position: relative;
  z-index: 1;
}

.bg-gradient {
  position: fixed;
  inset: 0;
  background:
    radial-gradient(ellipse 80% 60% at 50% 0%, #1a0b2e 0%, transparent 50%),
    radial-gradient(ellipse 80% 60% at 50% 100%, #0a0a1a 0%, transparent 50%),
    #060608;
  z-index: 0;
}

.bg-grid {
  position: fixed;
  inset: 0;
  background-image:
    linear-gradient(rgba(255,255,255,0.025) 1px, transparent 1px),
    linear-gradient(90deg, rgba(255,255,255,0.025) 1px, transparent 1px);
  background-size: 48px 48px;
  pointer-events: none;
  z-index: 0;
}

.bg-glow {
  position: fixed;
  border-radius: 50%;
  filter: blur(100px);
  pointer-events: none;
  z-index: 0;
  opacity: 0.35;
}

.bg-glow-1 {
  width: 500px; height: 500px;
  top: -200px; left: -200px;
  background: radial-gradient(circle, #7c6dfa 0%, transparent 70%);
  animation: drift1 20s ease-in-out infinite;
}

.bg-glow-2 {
  width: 500px; height: 500px;
  bottom: -200px; right: -200px;
  background: radial-gradient(circle, #fa6d9a 0%, transparent 70%);
  animation: drift2 24s ease-in-out infinite;
}

@keyframes drift1 {
  0%,100% { transform: translate(0,0) scale(1); }
  50% { transform: translate(80px, 60px) scale(1.15); }
}

@keyframes drift2 {
  0%,100% { transform: translate(0,0) scale(1); }
  50% { transform: translate(-80px, -60px) scale(1.2); }
}

.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 14px 20px;
  padding-top: calc(14px + env(safe-area-inset-top, 0px));
  border-bottom: 1px solid rgba(255,255,255,0.06);
  background: rgba(6,6,8,0.7);
  backdrop-filter: blur(20px) saturate(180%);
  -webkit-backdrop-filter: blur(20px) saturate(180%);
  position: sticky; top: 0; z-index: 20;
  flex-shrink: 0;
}

.header-left { display: flex; align-items: center; gap: 12px; }

.logo-mark {
  border-radius: 11px;
  background: linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%);
  display: flex; align-items: center; justify-content: center;
  font-family: 'Syne', sans-serif;
  font-weight: 800;
  color: white; letter-spacing: 0.5px;
  box-shadow: 0 4px 14px rgba(124,109,250,0.4), inset 0 1px 0 rgba(255,255,255,0.2);
  flex-shrink: 0;
}

.brand {
  font-family: 'Syne', sans-serif;
  font-weight: 700; font-size: 18px; letter-spacing: 0.06em;
  color: #ffffff;
}

.brand-sub {
  font-size: 10px; color: #5a5a70;
  text-transform: uppercase; letter-spacing: 0.12em;
  font-weight: 500; margin-top: 1px;
}

.header-right { display: flex; align-items: center; gap: 8px; }

.icon-ghost {
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 10px;
  width: 36px; height: 36px;
  font-size: 16px; color: #a0a0b8;
  cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  transition: all 0.15s ease;
}

.icon-ghost:hover:not(:disabled) {
  background: rgba(255,255,255,0.08);
  border-color: rgba(124,109,250,0.4);
  color: #fff;
}

.icon-ghost:disabled { opacity: 0.4; cursor: not-allowed; }

.tabs {
  display: flex;
  padding: 12px 16px 0;
  gap: 6px;
  flex-shrink: 0;
}

.tab {
  flex: 1;
  padding: 10px 14px;
  background: transparent;
  border: none;
  color: #6a6a80;
  font-size: 13px; font-weight: 500;
  cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  gap: 7px;
  border-radius: 10px;
  transition: color 0.2s, background 0.2s;
}

.tab.active {
  color: #fff;
  background: rgba(124,109,250,0.15);
}

.tab:hover:not(.active) {
  color: #a0a0b8;
  background: rgba(255,255,255,0.03);
}

.tab-icon { font-size: 14px; }

.main {
  flex: 1;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  padding-top: 8px;
}

.empty {
  flex: 1;
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  padding: 28px 20px;
  overflow-y: auto;
  gap: 20px;
}

.empty-orb {
  position: relative;
  width: 88px; height: 88px;
  display: flex; align-items: center; justify-content: center;
  margin-bottom: 8px;
}

.orb-glow {
  position: absolute;
  inset: -20px;
  background: radial-gradient(circle, rgba(124,109,250,0.6) 0%, transparent 70%);
  filter: blur(24px);
  animation: orb-pulse 3s ease-in-out infinite;
}

@keyframes orb-pulse {
  0%,100% { opacity: 0.6; transform: scale(1); }
  50% { opacity: 1; transform: scale(1.1); }
}

.orb-inner {
  position: relative;
  width: 68px; height: 68px;
  border-radius: 20px;
  background: linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%);
  display: flex; align-items: center; justify-content: center;
  font-size: 28px;
  box-shadow: 0 10px 40px rgba(124,109,250,0.4), inset 0 1px 0 rgba(255,255,255,0.2);
  animation: float 4s ease-in-out infinite;
}

@keyframes float {
  0%,100% { transform: translateY(0); }
  50% { transform: translateY(-8px); }
}

.suggestions {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 9px;
  width: 100%;
  max-width: 440px;
  margin-top: 12px;
}

.sug-btn {
  background: rgba(255,255,255,0.03);
  border: 1px solid rgba(255,255,255,0.06);
  border-radius: 12px;
  padding: 12px 13px;
  text-align: left;
  cursor: pointer;
  color: #d8d8e8;
  font-family: inherit; font-size: 12.5px;
  display: flex; align-items: flex-start;
  gap: 9px; line-height: 1.4;
  transition: all 0.18s ease;
}

.sug-btn:hover, .sug-btn:active {
  background: rgba(255,255,255,0.06);
  border-color: rgba(124,109,250,0.4);
  transform: translateY(-2px);
}

.sug-icon { font-size: 15px; flex-shrink: 0; margin-top: 1px; }

.messages {
  flex: 1;
  overflow-y: auto;
  padding: 4px 14px 8px;
  display: flex; flex-direction: column; gap: 14px;
  -webkit-overflow-scrolling: touch;
}

.msg-row {
  display: flex; gap: 9px; align-items: flex-end;
  animation: msg-in 0.3s cubic-bezier(0.2, 0.8, 0.2, 1);
}

@keyframes msg-in {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: translateY(0); }
}

.msg-row.user { flex-direction: row-reverse; }

.msg-row .logo-mark {
  width: 28px; height: 28px;
  font-size: 9px; border-radius: 50%;
}

.user-avatar {
  width: 28px; height: 28px; border-radius: 50%;
  background: rgba(255,255,255,0.06);
  border: 1px solid rgba(255,255,255,0.12);
  color: #8a8aa0;
  display: flex; align-items: center; justify-content: center;
  font-weight: 700; font-size: 11px;
  flex-shrink: 0;
  font-family: 'Syne', sans-serif;
}

.bubble {
  max-width: 78%;
  padding: 11px 15px;
  font-size: 14px; line-height: 1.6;
  border-radius: 18px;
}

.bubble.user {
  background: linear-gradient(135deg, #7c6dfa 0%, #6c5ce7 100%);
  color: white;
  border-bottom-right-radius: 5px;
  box-shadow: 0 4px 16px rgba(124,109,250,0.25);
}

.bubble.agent {
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  color: #e4e4f0;
  border-bottom-left-radius: 5px;
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
}

.bubble.error {
  background: rgba(250,109,154,0.1);
  border: 1px solid rgba(250,109,154,0.3);
  color: #fa6d9a;
}

.ts {
  font-size: 10px; margin-top: 5px; opacity: 0.4;
}

.bubble.user .ts { text-align: right; }
.bubble.agent .ts { text-align: left; }

.typing { display: flex; gap: 5px; padding: 2px 0; }

.typing span {
  width: 7px; height: 7px; border-radius: 50%;
  background: rgba(255,255,255,0.4);
  display: block;
  animation: bounce 1.3s ease infinite;
}

.typing span:nth-child(2) { animation-delay: 0.15s; }
.typing span:nth-child(3) { animation-delay: 0.3s; }

@keyframes bounce {
  0%,60%,100% { transform: translateY(0); opacity: 0.4; }
  30% { transform: translateY(-6px); opacity: 1; }
}

.msg-content strong { color: #fff; font-weight: 600; }
.msg-content em { font-style: italic; color: #c0c0dc; }
.ax-h1 { font-family: 'Syne', sans-serif; font-size: 17px; font-weight: 700; margin: 12px 0 5px; color: #fff; }
.ax-h2 { font-family: 'Syne', sans-serif; font-size: 15px; font-weight: 600; margin: 10px 0 4px; color: #fff; }
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
  font-family: 'SF Mono', 'Courier New', monospace;
  font-size: 12px; margin: 8px 0;
  color: #6dfacc; line-height: 1.5; white-space: pre;
}

.ax-inline {
  background: rgba(0,0,0,0.35);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 5px;
  padding: 1px 6px;
  font-family: 'SF Mono', monospace;
  font-size: 12.5px; color: #fa6d9a;
}

.composer {
  padding: 10px 14px;
  padding-bottom: calc(10px + env(safe-area-inset-bottom, 0px));
  flex-shrink: 0;
  position: relative;
}

.clear-chat-btn {
  position: absolute;
  top: -26px; left: 50%;
  transform: translateX(-50%);
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  color: #6a6a80;
  border-radius: 100px;
  padding: 4px 11px;
  font-size: 11px; cursor: pointer;
  display: flex; align-items: center; gap: 5px;
}

.input-wrap {
  display: flex; align-items: flex-end; gap: 8px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 18px;
  padding: 8px 8px 8px 16px;
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  transition: all 0.2s ease;
}

.input-wrap:focus-within {
  border-color: rgba(124,109,250,0.5);
  background: rgba(255,255,255,0.06);
}

.input-wrap textarea {
  flex: 1;
  background: transparent; border: none; outline: none;
  color: #e8e8f0;
  font-family: inherit; font-size: 15px;
  resize: none;
  min-height: 24px; max-height: 120px;
  line-height: 1.5;
  overflow-y: auto;
  padding: 5px 0;
}

.input-wrap textarea::placeholder { color: #4a4a5c; }

.send-btn {
  width: 36px; height: 36px;
  border-radius: 12px; border: none;
  background: linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%);
  color: white; cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
  font-size: 17px; font-weight: bold;
  transition: all 0.15s ease;
  box-shadow: 0 2px 8px rgba(124,109,250,0.3);
}

.send-btn:hover:not(:disabled) {
  transform: scale(1.06);
}

.send-btn:disabled { opacity: 0.25; cursor: not-allowed; box-shadow: none; }

.settings-overlay {
  position: fixed;
  inset: 0;
  background: rgba(6,6,8,0.85);
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  display: flex; align-items: center; justify-content: center;
  padding: 20px;
  z-index: 100;
}

.settings-card {
  width: 100%; max-width: 380px;
  background: rgba(20,20,30,0.9);
  border: 1px solid rgba(255,255,255,0.1);
  border-radius: 24px;
  padding: 36px 28px 28px;
  position: relative;
  overflow: hidden;
  box-shadow: 0 20px 60px rgba(0,0,0,0.5);
}

.settings-glow {
  position: absolute;
  top: -100px; left: 50%;
  transform: translateX(-50%);
  width: 300px; height: 300px;
  background: radial-gradient(circle, rgba(124,109,250,0.4) 0%, transparent 70%);
  filter: blur(40px);
  pointer-events: none;
}

.settings-header {
  text-align: center;
  margin-bottom: 28px;
  position: relative;
}

.settings-header .logo-mark {
  margin: 0 auto 12px;
  border-radius: 18px;
}

.settings-brand {
  font-family: 'Syne', sans-serif;
  font-weight: 700; font-size: 26px;
  letter-spacing: 0.04em;
  color: #ffffff;
}

.settings-sub {
  font-size: 11px; color: #5a5a70;
  text-transform: uppercase; letter-spacing: 0.15em;
  margin-top: 4px;
}

.settings-form {
  display: flex; flex-direction: column; gap: 14px;
  position: relative;
}

.field-label {
  font-size: 12px; color: #8a8aa0;
  font-weight: 500; margin-bottom: 2px;
}

.key-row { display: flex; gap: 8px; }

.key-row input {
  flex: 1;
  background: rgba(0,0,0,0.3);
  border: 1px solid rgba(255,255,255,0.1);
  border-radius: 12px;
  padding: 12px 14px;
  color: #e8e8f0;
  font-family: inherit; font-size: 14px;
  outline: none;
}

.key-row input:focus { border-color: rgba(124,109,250,0.5); }
.key-row input::placeholder { color: #4a4a5c; }

.eye-btn {
  background: rgba(0,0,0,0.3);
  border: 1px solid rgba(255,255,255,0.1);
  border-radius: 12px;
  width: 46px; height: 46px;
  cursor: pointer; font-size: 16px;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}

.cta {
  width: 100%;
  padding: 13px;
  background: linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%);
  border: none; border-radius: 12px;
  color: white;
  font-family: 'Syne', sans-serif;
  font-weight: 700; font-size: 15px;
  letter-spacing: 0.05em;
  cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  gap: 8px;
  box-shadow: 0 4px 14px rgba(124,109,250,0.3);
}

.cta:disabled { opacity: 0.4; cursor: not-allowed; box-shadow: none; }

.cta-arrow { transition: transform 0.2s ease; }

.cta-secondary {
  width: 100%;
  padding: 11px;
  background: transparent;
  border: 1px solid rgba(255,255,255,0.1);
  border-radius: 12px;
  color: #8a8aa0;
  font-family: inherit;
  font-weight: 500; font-size: 13px;
  cursor: pointer;
}

.get-key-link {
  color: #6a6a80;
  font-size: 12px;
  text-align: center;
  text-decoration: none;
  margin-top: 4px;
}

.error-box {
  background: rgba(250,109,154,0.08);
  border: 1px solid rgba(250,109,154,0.3);
  color: #fa6d9a;
  border-radius: 10px;
  padding: 10px 14px;
  font-size: 13px;
  line-height: 1.5;
  word-break: break-word;
}

.email-main {
  flex: 1;
  display: flex; flex-direction: column;
  overflow: hidden;
}

.account-bar {
  display: flex; align-items: center; justify-content: space-between;
  padding: 10px 16px 12px;
  margin: 0 14px;
  border-bottom: 1px solid rgba(255,255,255,0.05);
  flex-shrink: 0;
}

.account-info { display: flex; align-items: center; gap: 10px; min-width: 0; }

.account-avatar {
  width: 34px; height: 34px; border-radius: 50%;
  background: linear-gradient(135deg, #00b7c3 0%, #7c6dfa 100%);
  display: flex; align-items: center; justify-content: center;
  color: white; font-weight: 700; font-size: 14px;
  flex-shrink: 0;
}

.account-name {
  font-size: 13px; font-weight: 600; color: #e8e8f0;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  max-width: 180px;
}

.account-email {
  font-size: 11px; color: #6a6a80;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  max-width: 180px;
}

.account-actions { display: flex; gap: 6px; }

.account-actions .icon-ghost {
  width: 32px; height: 32px; font-size: 14px;
}

.email-loading {
  flex: 1;
  display: flex; flex-direction: column;
  align-items: center; justify-content: center;
  gap: 16px;
}

.loading-text { color: #6a6a80; font-size: 13px; }

.empty-inbox {
  flex: 1;
  display: flex; align-items: center; justify-content: center;
  color: #5a5a70; font-size: 14px;
}

.email-list {
  flex: 1;
  overflow-y: auto;
  padding: 12px 14px calc(14px + env(safe-area-inset-bottom, 0px));
  display: flex; flex-direction: column; gap: 10px;
  -webkit-overflow-scrolling: touch;
}

.email-card {
  background: rgba(255,255,255,0.03);
  border: 1px solid rgba(255,255,255,0.06);
  border-radius: 14px;
  padding: 13px 15px;
  position: relative;
}

.email-card.unread {
  background: rgba(124,109,250,0.06);
  border-color: rgba(124,109,250,0.2);
}

.email-card.unread::before {
  content: '';
  position: absolute;
  top: 17px; left: 5px;
  width: 4px; height: 4px; border-radius: 50%;
  background: #7c6dfa;
  box-shadow: 0 0 8px #7c6dfa;
}

.email-head {
  display: flex; justify-content: space-between; align-items: center;
  gap: 10px;
  margin-bottom: 4px;
}

.email-from {
  font-size: 13px; font-weight: 600; color: #e8e8f0;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  flex: 1;
}

.email-time {
  font-size: 11px; color: #5a5a70;
  flex-shrink: 0;
}

.email-subject {
  font-size: 13.5px; color: #c8c8dc;
  margin-bottom: 5px;
  line-height: 1.4;
  word-break: break-word;
}

.email-preview {
  font-size: 12px; color: #6a6a80;
  line-height: 1.5;
  margin-bottom: 10px;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

.email-actions {
  display: flex; justify-content: flex-end;
  margin-top: 2px;
}

.draft-btn {
  background: linear-gradient(135deg, #7c6dfa 0%, #fa6d9a 100%);
  border: none;
  border-radius: 10px;
  padding: 8px 14px;
  color: white;
  font-family: inherit;
  font-size: 12.5px; font-weight: 600;
  cursor: pointer;
  display: flex; align-items: center; gap: 6px;
  box-shadow: 0 2px 10px rgba(124,109,250,0.3);
}

.draft-btn:disabled { opacity: 0.7; cursor: default; }

.drafted-pill {
  background: rgba(109,250,204,0.1);
  border: 1px solid rgba(109,250,204,0.3);
  color: #6dfacc;
  border-radius: 10px;
  padding: 7px 12px;
  font-size: 12px; font-weight: 500;
}

.spinner {
  width: 12px; height: 12px;
  border: 2px solid rgba(255,255,255,0.3);
  border-top-color: white;
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
  display: inline-block;
}

@keyframes spin { to { transform: rotate(360deg); } }

::-webkit-scrollbar { width: 3px; }
::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); border-radius: 3px; }
::-webkit-scrollbar-track { background: transparent; }
`
