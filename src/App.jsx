import { useState, useRef, useEffect, useCallback } from 'react'

const SYSTEM_PROMPT = `You are AXIOM — an advanced, all-purpose AI agent built for Abdullah in Riyadh, Saudi Arabia.

You handle everything: daily tasks, research, writing, coding, analysis, planning, creative work, math, science, philosophy, business strategy, personal advice, and highly complex matters.

Key behaviors:
- Bilingual: reply in the same language the user writes in (Arabic or English)
- Saudi context aware: TASI, Tadawul, Vision 2030, local business environment
- Sharp and direct — never vague or overly cautious
- Structured for complex topics, concise for simple ones
- Use markdown formatting: **bold**, bullet points, code blocks, headers when it helps clarity
- Honest about knowledge cutoff (early 2025)

Always be the most useful assistant Abdullah has ever used.`

const SUGGESTED = [
  { icon: '📈', text: 'How do I analyze a TASI stock?' },
  { icon: '✍️', text: 'Write a professional email for me' },
  { icon: '💻', text: 'Help me write a Python script' },
  { icon: '🧠', text: 'Explain a complex topic simply' },
  { icon: '📊', text: 'Analyze pros and cons of my idea' },
  { icon: '🗓️', text: 'Help me plan my week' },
]

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

function TypingDots() {
  return (
    <div style={{display:'flex',gap:4,alignItems:'center',padding:'4px 2px'}}>
      {[0,1,2].map(i => (
        <span key={i} style={{
          width:6,height:6,borderRadius:'50%',background:'#5a5a70',display:'block',
          animation:`axbounce 1.2s ease infinite`,
          animationDelay:`${i*0.15}s`
        }}/>
      ))}
    </div>
  )
}

function Message({ msg }) {
  const isUser = msg.role === 'user'
  return (
    <div style={{
      display:'flex', gap:8, alignItems:'flex-end',
      flexDirection: isUser ? 'row-reverse' : 'row',
      animation:'axfade 0.2s ease',
    }}>
      <div style={{
        width:28, height:28, borderRadius:'50%', flexShrink:0,
        display:'flex', alignItems:'center', justifyContent:'center',
        fontFamily:"'Syne',sans-serif", fontWeight:800, fontSize:9,
        color:'white', letterSpacing:0.3,
        background: isUser ? '#18181f' : 'linear-gradient(135deg,#7c6dfa,#fa6d9a)',
        border: isUser ? '1px solid #ffffff18' : 'none',
        boxShadow: isUser ? 'none' : '0 0 8px #7c6dfa28',
      }}>
        {isUser ? 'U' : 'AX'}
      </div>
      <div style={{
        maxWidth:'80%', padding:'10px 14px', fontSize:14, lineHeight:1.65,
        borderRadius:16,
        borderBottomRightRadius: isUser ? 4 : 16,
        borderBottomLeftRadius:  isUser ? 16 : 4,
        background: isUser ? '#7c6dfa' : '#18181f',
        border: isUser ? 'none' : '1px solid #ffffff0f',
        color: isUser ? 'white' : '#e4e4f0',
        ...(msg.isError && { background:'#1f1215', border:'1px solid #fa6d9a40', color:'#fa6d9a' }),
      }}>
        {msg.loading
          ? <TypingDots/>
          : <div className="ax-content" dangerouslySetInnerHTML={{__html: parseMarkdown(msg.content)}}/>
        }
        {msg.timestamp && !msg.loading && (
          <div style={{fontSize:10, marginTop:4, opacity:0.35, textAlign: isUser ? 'right' : 'left'}}>
            {msg.timestamp}
          </div>
        )}
      </div>
    </div>
  )
}

function Settings({ onSave }) {
  const [key, setKey]         = useState(() => localStorage.getItem('axiom_api_key') || '')
  const [visible, setVisible] = useState(false)
  const save = () => {
    const k = key.trim()
    if (!k) return
    localStorage.setItem('axiom_api_key', k)
    onSave(k)
  }
  return (
    <div style={{height:'100%', display:'flex', alignItems:'center', justifyContent:'center', padding:24, background:'#09090f'}}>
      <div style={{width:'100%', maxWidth:360, background:'#111119', border:'1px solid #ffffff18', borderRadius:20, padding:'32px 24px', display:'flex', flexDirection:'column', alignItems:'center', gap:24}}>
        <div style={{textAlign:'center'}}>
          <div style={{width:56, height:56, borderRadius:14, margin:'0 auto 12px', background:'linear-gradient(135deg,#7c6dfa,#fa6d9a)', display:'flex', alignItems:'center', justifyContent:'center', fontFamily:"'Syne',sans-serif", fontWeight:800, fontSize:14, color:'white', boxShadow:'0 0 24px #7c6dfa30'}}>AX</div>
          <div style={{fontFamily:"'Syne',sans-serif", fontWeight:700, fontSize:22, background:'linear-gradient(90deg,#e4e4f0,#5a5a70)', WebkitBackgroundClip:'text', WebkitTextFillColor:'transparent'}}>AXIOM</div>
          <div style={{fontSize:10, color:'#5a5a70', textTransform:'uppercase', letterSpacing:'0.08em', marginTop:3}}>AI Agent · Powered by Claude</div>
        </div>
        <div style={{width:'100%', display:'flex', flexDirection:'column', gap:12}}>
          <p style={{fontSize:13, color:'#5a5a70', textAlign:'center', lineHeight:1.5}}>Enter your Anthropic API key to get started</p>
          <div style={{display:'flex', gap:8}}>
            <input type={visible ? 'text' : 'password'} placeholder="sk-ant-..." value={key} onChange={e => setKey(e.target.value)} onKeyDown={e => e.key === 'Enter' && save()} autoComplete="off" spellCheck={false} style={{flex:1, background:'#18181f', border:'1px solid #ffffff18', borderRadius:10, padding:'10px 12px', color:'#e4e4f0', fontFamily:"'DM Sans',sans-serif", fontSize:14, outline:'none'}}/>
            <button onClick={() => setVisible(v => !v)} style={{background:'#18181f', border:'1px solid #ffffff18', borderRadius:10, width:42, height:42, cursor:'pointer', fontSize:16, display:'flex', alignItems:'center', justifyContent:'center'}}>{visible ? '🙈' : '👁️'}</button>
          </div>
          <button onClick={save} disabled={!key.trim()} style={{width:'100%', padding:13, background:'linear-gradient(135deg,#7c6dfa,#fa6d9a)', border:'none', borderRadius:10, color:'white', fontFamily:"'Syne',sans-serif", fontWeight:700, fontSize:15, cursor:'pointer', opacity: key.trim() ? 1 : 0.4, letterSpacing:'0.03em'}}>Launch AXIOM →</button>
          <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noreferrer" style={{color:'#5a5a70', fontSize:12, textAlign:'center', textDecoration:'none'}}>Get your API key at console.anthropic.com</a>
        </div>
      </div>
    </div>
  )
}

export default function App() {
  const [apiKey, setApiKey]     = useState(() => localStorage.getItem('axiom_api_key') || '')
  const [showSettings, setShow] = useState(false)
  const [messages, setMessages] = useState([])
  const [history, setHistory]   = useState([])
  const [input, setInput]       = useState('')
  const [loading, setLoading]   = useState(false)
  const bottomRef = useRef(null)
  const taRef     = useRef(null)

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: 'smooth' }) }, [messages])

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
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model: 'claude-haiku-4-5-20251001',
          max_tokens: 1024,
          system: SYSTEM_PROMPT,
          messages: newHistory,
        }),
      })
      const data = await res.json().catch(() => null)
      if (!res.ok) throw new Error(data?.error?.message || `Error ${res.status}`)
      if (!data)   throw new Error('No response from server.')
      const reply = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n\n').trim()
      if (!reply) throw new Error('Empty response. Please try again.')
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

  if (!apiKey || showSettings) return <Settings onSave={k => { setApiKey(k); setShow(false) }}/>

  return (
    <div style={{display:'flex',flexDirection:'column',height:'100%',maxWidth:680,margin:'0 auto',position:'relative',zIndex:1}}>
      <header style={{display:'flex', alignItems:'center', justifyContent:'space-between', padding:'12px 16px', paddingTop:`calc(12px + env(safe-area-inset-top, 0px))`, borderBottom:'1px solid #ffffff0f', background:'rgba(9,9,15,0.97)', backdropFilter:'blur(12px)', flexShrink:0, position:'sticky', top:0, zIndex:20}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <div style={{width:34, height:34, borderRadius:9, background:'linear-gradient(135deg,#7c6dfa,#fa6d9a)', display:'flex', alignItems:'center', justifyContent:'center', fontFamily:"'Syne',sans-serif", fontWeight:800, fontSize:10, color:'white', boxShadow:'0 0 14px #7c6dfa30', flexShrink:0}}>AX</div>
          <div>
            <div style={{fontFamily:"'Syne',sans-serif", fontWeight:700, fontSize:16, background:'linear-gradient(90deg,#e4e4f0,#5a5a70)', WebkitBackgroundClip:'text', WebkitTextFillColor:'transparent'}}>AXIOM</div>
            <div style={{fontSize:9,color:'#5a5a70',textTransform:'uppercase',letterSpacing:'0.08em'}}>AI Agent</div>
          </div>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:6}}>
          <div style={{display:'flex', alignItems:'center', gap:6, background:'#18181f', border:'1px solid #ffffff18', borderRadius:100, padding:'4px 10px', fontSize:11, color:'#5a5a70'}}>
            <span style={{width:6, height:6, borderRadius:'50%', background:'#6dfacc', boxShadow:'0 0 6px #6dfacc', animation:'axblink 2s ease-in-out infinite', display:'block'}}/>
            {loading ? 'Thinking…' : 'Ready'}
          </div>
          <button onClick={() => setShow(true)} style={iconBtn}>⚙️</button>
          {!isEmpty && <button onClick={clear} style={iconBtn}>🗑️</button>}
        </div>
      </header>
      {isEmpty ? (
        <div style={{flex:1, display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', padding:'24px 16px', gap:24, overflowY:'auto'}}>
          <div style={{textAlign:'center'}}>
            <div style={{width:58, height:58, borderRadius:16, background:'linear-gradient(135deg,#7c6dfa,#fa6d9a)', margin:'0 auto 14px', display:'flex', alignItems:'center', justifyContent:'center', fontSize:24, boxShadow:'0 0 28px #7c6dfa22', animation:'axfloat 3.5s ease-in-out infinite'}}>⚡</div>
            <div style={{fontFamily:"'Syne',sans-serif", fontWeight:800, fontSize:22, background:'linear-gradient(135deg,#e4e4f0 30%,#5a5a70)', WebkitBackgroundClip:'text', WebkitTextFillColor:'transparent', marginBottom:6}}>What can I do for you?</div>
            <div style={{color:'#5a5a70',fontSize:13,lineHeight:1.6,maxWidth:260,margin:'0 auto'}}>Daily tasks, research, coding, analysis — ask me anything.</div>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(2,1fr)',gap:8,width:'100%',maxWidth:420}}>
            {SUGGESTED.map((s,i) => (
              <button key={i} onClick={() => send(s.text)} style={{background:'#111119', border:'1px solid #ffffff0f', borderRadius:10, padding:'10px 12px', textAlign:'left', cursor:'pointer', color:'#e4e4f0', fontFamily:"'DM Sans',sans-serif", fontSize:12, display:'flex', alignItems:'flex-start', gap:7, lineHeight:1.35}}>
                <span style={{fontSize:14,flexShrink:0,marginTop:1}}>{s.icon}</span>
                <span>{s.text}</span>
              </button>
            ))}
          </div>
        </div>
      ) : (
        <div style={{flex:1, overflowY:'auto', padding:'16px 12px', display:'flex', flexDirection:'column', gap:12, WebkitOverflowScrolling:'touch'}}>
          {messages.map((msg,i) => <Message key={i} msg={msg}/>)}
          <div ref={bottomRef}/>
        </div>
      )}
      <div style={{padding:'10px 12px', paddingBottom:`calc(10px + env(safe-area-inset-bottom, 0px))`, borderTop:'1px solid #ffffff0f', background:'rgba(9,9,15,0.98)', flexShrink:0}}>
        <div style={{display:'flex', alignItems:'flex-end', gap:8, background:'#111119', border:'1px solid #ffffff18', borderRadius:14, padding:'8px 9px'}}>
          <textarea ref={taRef} rows={1} value={input} onChange={e => { setInput(e.target.value); e.target.style.height = 'auto'; e.target.style.height = Math.min(e.target.scrollHeight, 120) + 'px' }} onKeyDown={handleKey} placeholder="Ask me anything…" disabled={loading} style={{flex:1, background:'transparent', border:'none', outline:'none', color:'#e4e4f0', fontFamily:"'DM Sans',sans-serif", fontSize:15, resize:'none', minHeight:24, maxHeight:120, lineHeight:1.5, overflowY:'auto'}}/>
          <button onClick={() => send()} disabled={loading || !input.trim()} style={{width:34, height:34, borderRadius:9, border:'none', background:'linear-gradient(135deg,#7c6dfa,#fa6d9a)', color:'white', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0, fontSize:16, fontWeight:'bold', opacity: (loading || !input.trim()) ? 0.3 : 1}}>↑</button>
        </div>
        <div style={{textAlign:'center',fontSize:10,color:'#5a5a70',marginTop:5,opacity:0.5}}>Enter to send · Shift+Enter for new line</div>
      </div>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500&display=swap');
        *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
        html,body{height:100%;overflow:hidden;background:#09090f;-webkit-font-smoothing:antialiased}
        #root{height:100%;display:flex;flex-direction:column}
        @keyframes axbounce{0%,60%,100%{transform:translateY(0)}30%{transform:translateY(-5px)}}
        @keyframes axfade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
        @keyframes axfloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-7px)}}
        @keyframes axblink{0%,100%{opacity:1}50%{opacity:.3}}
        .ax-content strong{color:#fff;font-weight:600}
        .ax-content em{font-style:italic;color:#c0c0dc}
        .ax-h1{font-family:'Syne',sans-serif;font-size:16px;font-weight:700;margin:12px 0 5px}
        .ax-h2{font-family:'Syne',sans-serif;font-size:15px;font-weight:600;margin:10px 0 4px}
        .ax-h3{font-size:14px;font-weight:600;margin:8px 0 3px}
        .ax-li{padding:2px 0 2px 14px;position:relative;list-style:none;display:block}
        .ax-li::before{content:'›';position:absolute;left:0;color:#7c6dfa;font-weight:700}
        .ax-li-n{padding:2px 0 2px 4px;display:block}
        .ax-gap{height:8px}
        .ax-code{background:#07070e;border:1px solid #ffffff18;border-radius:8px;padding:10px;overflow-x:auto;font-family:'Courier New',monospace;font-size:12px;margin:8px 0;color:#6dfacc;line-height:1.5;white-space:pre}
        .ax-inline{background:#07070e;border:1px solid #ffffff18;border-radius:4px;padding:1px 5px;font-family:monospace;font-size:12px;color:#fa6d9a}
        textarea::placeholder{color:#5a5a70}
        input::placeholder{color:#5a5a70}
        ::-webkit-scrollbar{width:3px}
        ::-webkit-scrollbar-thumb{background:#ffffff18;border-radius:3px}
      `}</style>
    </div>
  )
}

const iconBtn = {
  background:'transparent', border:'1px solid #ffffff18',
  borderRadius:8, width:30, height:30, fontSize:14,
  cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center',
}
