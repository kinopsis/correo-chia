import { corsHeaders } from '../_shared/cors.ts'
import { supabase } from '../_shared/supabase.ts'

const GRAPH_API = 'https://graph.microsoft.com/v1.0'

async function readSettings(): Promise<Record<string, unknown>> {
  const { data } = await supabase.from('email_settings').select('*').eq('id', 1).single()
  return (data as Record<string, unknown>) || {}
}

async function sendViaSmtp(settings: Record<string, unknown>, email: {
  from: string; fromName: string; to: string; cc?: string[];
  subject: string; html: string; text: string;
  inReplyTo?: string; references?: string;
}): Promise<void> {
  const host = (settings.smtp_host as string) || Deno.env.get('SMTP_HOST') || 'smtp.office365.com'
  const port = (settings.smtp_port as number) || 587
  const user = (settings.smtp_user as string) || Deno.env.get('SMTP_USER') || email.from
  const pass = Deno.env.get('SMTP_PASSWORD')
  if (!pass) throw new Error('SMTP_PASSWORD no configurado en secrets de Supabase')

  const conn = await Deno.connectTls({ hostname: host, port })
  const buf = new TextEncoder(); const dec = new TextDecoder()

  async function cmd(command: string, expectCode = 250): Promise<string> {
    if (command) await conn.write(buf.encode(command + '\r\n'))
    const chunk = new Uint8Array(4096); const n = await conn.read(chunk)
    const response = dec.decode(chunk.subarray(0, n ?? 0)); const code = parseInt(response.substring(0, 3))
    if (code !== expectCode && expectCode !== 0) {
      const cn = command.split(' ')[0]
      if (cn === 'AUTH' && code === 535) throw new Error('Error de autenticacion SMTP: credenciales rechazadas. Verifique SMTP_PASSWORD en secrets de Supabase.')
      throw new Error(`SMTP ${cn}: esperado ${expectCode}, recibido ${code}: ${response.slice(0, 200)}`)
    }
    return response
  }

  try {
    await cmd('', 220); await cmd(`EHLO ${host}`); await cmd('STARTTLS', 220)
    const tc = await Deno.startTls(conn, { hostname: host }); const tb = new TextEncoder(); const td = new TextDecoder()
    async function tlsCmd(command: string, expectCode = 250): Promise<string> {
      if (command) await tc.write(tb.encode(command + '\r\n'))
      const chunk = new Uint8Array(4096); const n = await tc.read(chunk)
      const response = td.decode(chunk.subarray(0, n ?? 0)); const code = parseInt(response.substring(0, 3))
      if (code !== expectCode && expectCode !== 0) {
        const cn = command.split(' ')[0]
        if (cn === 'AUTH' && code === 535) throw new Error('Error de autenticacion SMTP: credenciales rechazadas.')
        throw new Error(`SMTP TLS ${cn}: esperado ${expectCode}, recibido ${code}`)
      }
      return response
    }
    await tlsCmd(`EHLO ${host}`); await tlsCmd('AUTH LOGIN', 334); await tlsCmd(btoa(user), 334); await tlsCmd(btoa(pass), 235)
    await tlsCmd(`MAIL FROM:<${email.from}>`); await tlsCmd(`RCPT TO:<${email.to}>`); await tlsCmd('DATA', 354)
    const mid = `<${crypto.randomUUID()}@chia.gov.co>`
    let mime = `From: "${email.fromName}" <${email.from}>\r\nTo: <${email.to}>\r\n`
    if (email.cc?.length) mime += `Cc: ${email.cc.join(', ')}\r\n`
    mime += `Subject: ${email.subject}\r\nMessage-ID: ${mid}\r\n`
    if (email.inReplyTo) mime += `In-Reply-To: ${email.inReplyTo}\r\n`
    if (email.references) mime += `References: ${email.references} ${email.inReplyTo||''}\r\n`.trim() + '\r\n'
    mime += `MIME-Version: 1.0\r\nContent-Type: text/html; charset=UTF-8\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n`
    mime += email.html || email.text; mime += '\r\n.'
    await tlsCmd(mime); await tlsCmd('QUIT', 221); tc.close()
  } finally { try { conn.close() } catch { } }
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders })
  try {
    const { case_id, body_html, body_text, sent_by_user, attachment_ids, cc } = await req.json()
    if (!case_id || !sent_by_user) return new Response(JSON.stringify({ error: 'case_id y sent_by_user requeridos' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })

    const settings = await readSettings()
    const provider = (settings.provider as string) || 'smtp'
    const mailbox = (settings.mailbox as string) || Deno.env.get('MAILBOX') || 'contactenos@chia.gov.co'
    const fromName = (settings.from_name as string) || Deno.env.get('FROM_NAME') || 'Linea de Atencion Ciudadana - Alcaldia de Chia'

    const { data: caseData, error: caseError } = await supabase.from('cases').select('*, messages!inner(*)').eq('id', case_id).eq('messages.is_first_message', true).single()
    if (caseError || !caseData) return new Response(JSON.stringify({ error: 'Caso no encontrado' }), { status: 404, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })

    const originalMsg = caseData.messages?.[0]
    const subject = caseData.subject?.startsWith('RE:') ? caseData.subject : `RE: ${caseData.subject}`
    const messageId = `<${crypto.randomUUID()}@chia.gov.co>`

    let inReplyTo: string | undefined; let refs: string | undefined
    if (originalMsg?.message_id || originalMsg?.references) {
      inReplyTo = originalMsg.message_id
      refs = originalMsg.references ? `${originalMsg.references} ${originalMsg.message_id||''}`.trim() : originalMsg.message_id
    }

    if (provider === 'smtp') {
      await sendViaSmtp(settings, { from: mailbox, fromName, to: caseData.citizen_email, cc: cc?.length ? cc : undefined, subject, html: body_html || '', text: body_text || '', inReplyTo, references: refs })
    } else {
      const token = await getGraphToken()
      const sr = await fetch(`${GRAPH_API}/users/${mailbox}/sendMail`, { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ message: { subject, body: { contentType: 'HTML', content: body_html||body_text||'' }, from: { emailAddress: { address: mailbox, name: fromName } }, toRecipients: [{ emailAddress: { address: caseData.citizen_email } }] }, saveToSentItems: true }) })
      if (!sr.ok) { const et = await sr.text(); throw new Error(`Send failed: ${sr.status} - ${et}`) }
    }

    const { data: outMsg, error: mErr } = await supabase.from('messages').insert({
      case_id, message_id: messageId, in_reply_to: originalMsg?.message_id||null,
      references: originalMsg?.references ? `${originalMsg.references} ${originalMsg.message_id||''}`.trim() : originalMsg?.message_id||null,
      direction: 'outbound', from_addr: mailbox, from_name: fromName, to_addr: caseData.citizen_email, cc_addrs: cc||[],
      subject, body_text: body_text||'', body_html: body_html||'', has_attachments: attachment_ids?.length>0,
      sent_by_user, sent_at: new Date().toISOString(), created_at: new Date().toISOString()
    }).select('id').single()
    if (mErr) console.error('Error saving msg:', mErr)

    if (!caseData.first_reply_at) { await supabase.from('cases').update({ first_reply_at: new Date().toISOString(), status: 'in_progress', updated_at: new Date().toISOString() }).eq('id', case_id) }
    else { await supabase.from('cases').update({ updated_at: new Date().toISOString() }).eq('id', case_id) }

    return new Response(JSON.stringify({ success: true, message_id: outMsg?.id }), { headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
  } catch (error) {
    console.error('SMTP Sender error:', error)
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : 'Unknown error' }), { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
  }
})

async function getGraphToken(): Promise<string> {
  const tid = Deno.env.get('AZURE_TENANT_ID'); const cid = Deno.env.get('AZURE_CLIENT_ID'); const cs = Deno.env.get('AZURE_CLIENT_SECRET')
  if (!tid || !cid || !cs) throw new Error('Credenciales Azure AD no configuradas.')
  const resp = await fetch(`https://login.microsoftonline.com/${tid}/oauth2/v2.0/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: cid, client_secret: cs, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }) })
  if (!resp.ok) throw new Error(`Auth failed: ${resp.status}`)
  const data = await resp.json(); return data.access_token
}
