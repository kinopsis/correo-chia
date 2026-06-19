import { corsHeaders } from '../_shared/cors.ts'
import { supabase } from '../_shared/supabase.ts'
import { classify } from '../_shared/classification.ts'
import { assignCase } from '../_shared/assignment-engine.ts'

const GRAPH_API = 'https://graph.microsoft.com/v1.0'

async function readSettings(): Promise<Record<string, unknown>> {
  const { data } = await supabase.from('email_settings').select('*').eq('id', 1).single()
  return (data as Record<string, unknown>) || {}
}

interface ImapMessage {
  id: string; messageId: string; inReplyTo?: string; references?: string
  subject: string; fromAddr: string; fromName: string; toAddr: string
  bodyText: string; bodyHtml: string; hasAttachments: boolean; receivedDateTime: string
}

async function fetchViaImap(settings: Record<string, unknown>): Promise<ImapMessage[]> {
  const host = (settings.imap_host as string) || Deno.env.get('IMAP_HOST') || 'outlook.office365.com'
  const port = (settings.imap_port as number) || 993
  const user = (settings.imap_user as string) || Deno.env.get('IMAP_USER') || (settings.mailbox as string)
  const pass = Deno.env.get('IMAP_PASSWORD') || Deno.env.get('SMTP_PASSWORD')
  if (!user || !pass) throw new Error('Credenciales IMAP no configuradas. Configure IMAP_USER/IMAP_PASSWORD o SMTP_PASSWORD en secrets.')

  const esc = (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')
  const safeUser = esc(user); const safePass = esc(pass)

  const conn = await Deno.connectTls({ hostname: host, port })
  const encoder = new TextEncoder(); const decoder = new TextDecoder()
  let tagCounter = 0

  async function imapCmd(command: string, expectOk = true): Promise<string> {
    const tag = `A${String(++tagCounter).padStart(4, '0')}`
    await conn.write(encoder.encode(`${tag} ${command}\r\n`))
    const chunks: Uint8Array[] = []; let response = ''; let done = false
    while (!done) {
      const buf = new Uint8Array(4096); const n = await conn.read(buf)
      if (n === null) break
      const part = decoder.decode(buf.subarray(0, n))
      chunks.push(buf.subarray(0, n)); response += part
      if (new RegExp(`${tag} (OK|NO|BAD)`, 'm').test(response)) done = true
    }
    if (expectOk && !response.includes(`${tag} OK`)) {
      const cmdName = command.split(' ')[0]
      const isBad = response.includes(`${tag} BAD`)
      const ed = response.slice(Math.max(0, response.indexOf(tag)), Math.min(response.length, response.indexOf(tag) + 200)).trim()
      if (cmdName === 'LOGIN') throw new Error('Error de autenticacion IMAP: credenciales incorrectas o usuario no encontrado. Verifique el usuario y contrasena IMAP en Configuracion de Correo.')
      if (cmdName === 'SELECT') throw new Error('Error IMAP: no se pudo acceder a la bandeja INBOX. El buzon puede no existir o permisos IMAP no habilitados.')
      if (isBad) throw new Error(`Error de protocolo IMAP en ${cmdName}: servidor rechazo solicitud. Detalle: ${ed}`)
      throw new Error(`Error IMAP en ${cmdName}: servidor respondio con error. Detalle: ${ed}`)
    }
    return response
  }

  try {
    const greeting = new Uint8Array(4096); await conn.read(greeting)
    await imapCmd(`LOGIN "${safeUser}" "${safePass}"`); await imapCmd('SELECT INBOX')
    const sr = await imapCmd('SEARCH UNSEEN'); const um = sr.match(/\* SEARCH ([\d\s]+)/)
    if (!um || !um[1].trim()) return []
    const uids = um[1].trim().split(/\s+/); const msgs: ImapMessage[] = []
    for (const uid of uids.slice(0, 20)) {
      try {
        const fr = await imapCmd(`FETCH ${uid} (BODY[] INTERNALDATE)`)
        const bm = fr.match(/\* \d+ FETCH \(.*?BODY\[\] \{(\d+)\}\r\n([\s\S]*?)\r\n\)/)
        if (!bm) continue
        const raw = bm[2]; const dm = fr.match(/INTERNALDATE "([^"]+)"/)
        const he = raw.indexOf('\r\n\r\n'); const hs = he > 0 ? raw.substring(0, he) : raw
        const bs = he > 0 ? raw.substring(he + 4) : ''
        const gh = (n: string) => { const r = new RegExp(`^${n}:\\s*(.+)$`, 'im'); const m = hs.match(r); return m ? m[1].trim() : undefined }
        const frRaw = gh('From') || ''; const fm = frRaw.match(/(?:\"?([^\"]*)\"?\s*)?<?([^>]+)>?/)
        msgs.push({ id: uid, messageId: gh('Message-ID') || `<${crypto.randomUUID()}@imap.local>`, inReplyTo: gh('In-Reply-To'), references: gh('References'), subject: gh('Subject') || '(sin asunto)', fromAddr: fm?.[2]?.trim() || frRaw, fromName: fm?.[1]?.trim() || fm?.[2]?.trim() || '', toAddr: gh('To') || '', bodyText: bs.replace(/<[^>]+>/g, '').trim(), bodyHtml: bs, hasAttachments: hs.includes('Content-Type: multipart/mixed'), receivedDateTime: dm?.[1] || new Date().toISOString() })
      } catch (e) { console.error(`Error IMAP msg ${uid}:`, e) }
    }
    await imapCmd('LOGOUT', false); return msgs
  } finally { try { conn.close() } catch { } }
}

async function getGraphToken(): Promise<string> {
  const tid = Deno.env.get('AZURE_TENANT_ID'); const cid = Deno.env.get('AZURE_CLIENT_ID'); const cs = Deno.env.get('AZURE_CLIENT_SECRET')
  if (!tid || !cid || !cs) {
    const m = []; if (!tid) m.push('AZURE_TENANT_ID'); if (!cid) m.push('AZURE_CLIENT_ID'); if (!cs) m.push('AZURE_CLIENT_SECRET')
    throw new Error(`Credenciales Azure AD incompletas. Faltan: ${m.join(', ')}. Configure secrets en Supabase Dashboard > Edge Functions > imap-poller > Secrets.`)
  }
  const resp = await fetch(`https://login.microsoftonline.com/${tid}/oauth2/v2.0/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: cid, client_secret: cs, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }) })
  if (!resp.ok) throw new Error(`Auth failed: ${resp.status}`)
  const data = await resp.json(); return data.access_token
}

async function fetchUnreadMessages(token: string, mailbox: string) {
  const resp = await fetch(`${GRAPH_API}/users/${mailbox}/messages?$filter=isRead eq false&$top=50&$orderby=receivedDateTime desc`, { headers: { Authorization: `Bearer ${token}` } })
  if (!resp.ok) throw new Error(`Fetch messages failed: ${resp.status}`)
  const data = await resp.json(); return data.value || []
}

async function markAsRead(token: string, messageId: string, mailbox: string) {
  await fetch(`${GRAPH_API}/users/${mailbox}/messages/${messageId}`, { method: 'PATCH', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ isRead: true }) })
}

async function findExistingCase(msg: { inReplyTo?: string; references?: string; subject: string; from?: { emailAddress?: { address?: string } } }) {
  if (msg.inReplyTo) { const { data } = await supabase.from('messages').select('case_id').eq('message_id', msg.inReplyTo).single(); if (data) return data.case_id }
  if (msg.references) { for (const ref of msg.references.split(/[,\s]+/).filter(Boolean)) { const { data } = await supabase.from('messages').select('case_id').eq('message_id', ref.trim()).single(); if (data) return data.case_id } }
  return null
}

/// <reference types="https://deno.land/x/deno_types/deno.d.ts" />

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders })
  try {
    const settings = await readSettings()
    const provider = (settings.provider as string) || 'smtp'
    const mailbox = (settings.mailbox as string) || 'contactenos@chia.gov.co'

    if (provider === 'smtp') {
      const imapMessages = await fetchViaImap(settings)
      for (const msg of imapMessages) {
        const eci = await findExistingCase({ inReplyTo: msg.inReplyTo, references: msg.references, subject: msg.subject, from: { emailAddress: { address: msg.fromAddr } } })
        if (eci) {
          await supabase.from('messages').insert({ case_id: eci, message_id: msg.messageId, in_reply_to: msg.inReplyTo||null, direction: 'inbound', from_addr: msg.fromAddr, from_name: msg.fromName, to_addr: mailbox, subject: msg.subject, body_text: msg.bodyText, body_html: msg.bodyHtml, has_attachments: msg.hasAttachments, is_first_message: false, sent_at: msg.receivedDateTime, created_at: new Date().toISOString() })
          const { data: cc } = await supabase.from('cases').select('status').eq('id', eci).single()
          if (cc?.status === 'pending_citizen') await supabase.from('cases').update({ status: 'in_progress', updated_at: new Date().toISOString() }).eq('id', eci)
        } else {
          const subj = msg.subject || 'Sin asunto'; const bdy = msg.bodyText || msg.bodyHtml || ''
          const { category_id, priority } = await classify(subj, bdy)
          const { data: sr } = await supabase.from('sla_rules').select('max_response_h').eq('category_id', category_id).eq('priority', priority).eq('active', true).single()
          const rh = sr?.max_response_h ?? 8; const sdl = new Date(Date.now()+rh*60*60*1000).toISOString()
          const { data: nc, error: ce } = await supabase.from('cases').insert({ subject: subj, category_id, priority, status: 'new', citizen_email: msg.fromAddr, citizen_name: msg.fromName, first_message_id: msg.messageId, sla_deadline: sdl, created_at: new Date().toISOString(), updated_at: new Date().toISOString() }).select('id, case_number').single()
          if (ce || !nc) { console.error('Error creating case:', ce); continue }
          await supabase.from('messages').insert({ case_id: nc.id, message_id: msg.messageId, in_reply_to: msg.inReplyTo||null, direction: 'inbound', from_addr: msg.fromAddr, from_name: msg.fromName, to_addr: mailbox, subject: subj, body_text: bdy, body_html: msg.bodyHtml, has_attachments: msg.hasAttachments, is_first_message: true, sent_at: msg.receivedDateTime, created_at: new Date().toISOString() })
          await assignCase(nc.id, category_id)
          try { await fetch(`${Deno.env.get('SUPABASE_URL')}/functions/v1/ai-assistant`, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')}` }, body: JSON.stringify({ case_id: nc.id, subject: subj, body_text: bdy }) }) } catch (ae) { console.error('AI error:', ae) }
          console.log(`Caso ${nc.case_number} creado via IMAP`)
        }
      }
      return new Response(JSON.stringify({ success: true, processed: imapMessages.length, provider: 'imap' }), { headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
    }

    const token = await getGraphToken(); const messages = await fetchUnreadMessages(token, mailbox)
    for (const msg of messages) {
      const eci = await findExistingCase({ inReplyTo: msg.inferenceClassification !== 'focused' ? msg.inReplyTo : undefined, references: msg.references, subject: msg.subject || '', from: msg.from })
      if (eci) {
        await supabase.from('messages').insert({ case_id: eci, message_id: msg.id, in_reply_to: msg.inReplyTo||null, direction: 'inbound', from_addr: msg.from?.emailAddress?.address||'', from_name: msg.from?.emailAddress?.name||'', to_addr: mailbox, subject: msg.subject||'', body_text: msg.bodyPreview||'', body_html: msg.body?.content||'', has_attachments: msg.hasAttachments||false, is_first_message: false, imap_uid: null, sent_at: msg.receivedDateTime, created_at: new Date().toISOString() })
        const { data: cc } = await supabase.from('cases').select('status').eq('id', eci).single()
        if (cc?.status === 'pending_citizen') await supabase.from('cases').update({ status: 'in_progress', updated_at: new Date().toISOString() }).eq('id', eci)
      } else {
        const subj = msg.subject || 'Sin asunto'; const bdy = msg.bodyPreview || msg.body?.content || ''
        const { category_id, priority } = await classify(subj, bdy)
        const { data: sr } = await supabase.from('sla_rules').select('max_response_h').eq('category_id', category_id).eq('priority', priority).eq('active', true).single()
        const rh = sr?.max_response_h ?? 8; const sdl = new Date(Date.now()+rh*60*60*1000).toISOString()
        const { data: nc, error: ce } = await supabase.from('cases').insert({ subject: subj, category_id, priority, status: 'new', citizen_email: msg.from?.emailAddress?.address||'', citizen_name: msg.from?.emailAddress?.name||'', first_message_id: msg.id, sla_deadline: sdl, created_at: new Date().toISOString(), updated_at: new Date().toISOString() }).select('id, case_number').single()
        if (ce || !nc) { console.error('Error creating case:', ce); continue }
        await supabase.from('messages').insert({ case_id: nc.id, message_id: msg.id, in_reply_to: msg.inReplyTo||null, direction: 'inbound', from_addr: msg.from?.emailAddress?.address||'', from_name: msg.from?.emailAddress?.name||'', to_addr: mailbox, subject: subj, body_text: bdy, body_html: msg.body?.content||'', has_attachments: msg.hasAttachments||false, is_first_message: true, imap_uid: null, sent_at: msg.receivedDateTime, created_at: new Date().toISOString() })
        await assignCase(nc.id, category_id)
        try { await fetch(`${Deno.env.get('SUPABASE_URL')}/functions/v1/ai-assistant`, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')}` }, body: JSON.stringify({ case_id: nc.id, subject: subj, body_text: bdy }) }) } catch (ae) { console.error('AI error:', ae) }
        console.log(`Caso ${nc.case_number} creado y asignado`)
      }
      await markAsRead(token, msg.id, mailbox)
    }
    return new Response(JSON.stringify({ success: true, processed: messages.length, provider: 'graph_api' }), { headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
  } catch (error) {
    console.error('IMAP Poller error:', error)
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : 'Unknown error' }), { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
  }
})
