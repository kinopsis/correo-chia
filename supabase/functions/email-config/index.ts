import { corsHeaders } from '../_shared/cors.ts'
import { supabase } from '../_shared/supabase.ts'

/// <reference types="https://deno.land/x/deno_types/deno.d.ts" />

interface EmailStatus {
  provider: string; mailbox: string; from_name: string
  graph_configured: boolean; smtp_configured: boolean; imap_configured: boolean
}

async function readSettings(): Promise<Record<string, unknown>> {
  const { data, error } = await supabase.from('email_settings').select('*').eq('id', 1).single()
  if (error || !data) return {}
  return data as Record<string, unknown>
}

function buildStatus(settings: Record<string, unknown>): EmailStatus {
  const mailbox = (settings.mailbox as string) || Deno.env.get('MAILBOX') || 'contactenos@chia.gov.co'
  const fromName = (settings.from_name as string) || Deno.env.get('FROM_NAME') || 'Linea de Atencion Ciudadana - Alcaldia de Chia'
  const tid = Deno.env.get('AZURE_TENANT_ID'); const cid = Deno.env.get('AZURE_CLIENT_ID'); const cs = Deno.env.get('AZURE_CLIENT_SECRET')
  const graphOk = !!(tid && cid && cs)
  const sh = (settings.smtp_host as string) || Deno.env.get('SMTP_HOST'); const su = (settings.smtp_user as string) || Deno.env.get('SMTP_USER')
  const sp = Deno.env.get('SMTP_PASSWORD'); const smtpOk = !!(sh && su && sp)
  const ih = (settings.imap_host as string) || Deno.env.get('IMAP_HOST'); const iu = (settings.imap_user as string) || Deno.env.get('IMAP_USER')
  const ip = Deno.env.get('IMAP_PASSWORD'); const imapOk = !!(ih && iu && ip)
  const prov = (settings.provider as string) || 'none'
  const ep = prov !== 'none' ? prov : smtpOk ? 'smtp' : graphOk ? 'graph_api' : 'none'
  return { provider: ep, mailbox, from_name: fromName, graph_configured: graphOk, smtp_configured: smtpOk, imap_configured: imapOk }
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders })
  try {
    const body = await req.json(); const action = body.action as string

    if (action === 'get-status') {
      const s = await readSettings(); return new Response(JSON.stringify(buildStatus(s)), { headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
    }

    if (action === 'test') {
      const settings = await readSettings()
      const provider = body.provider as string || (settings.provider as string) || 'smtp'
      const mailbox = (body.mailbox as string) || (settings.mailbox as string) || Deno.env.get('MAILBOX') || 'contactenos@chia.gov.co'

      if (provider === 'graph_api') {
        const tid = Deno.env.get('AZURE_TENANT_ID'); const cid = Deno.env.get('AZURE_CLIENT_ID'); const cs = Deno.env.get('AZURE_CLIENT_SECRET')
        if (!tid || !cid || !cs) return new Response(JSON.stringify({ success: false, error: 'Credenciales Azure AD no configuradas.' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
        const resp = await fetch(`https://login.microsoftonline.com/${tid}/oauth2/v2.0/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: cid, client_secret: cs, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }) })
        if (!resp.ok) { const e = await resp.text(); return new Response(JSON.stringify({ success: false, error: `Error Azure AD: ${e.slice(0, 200)}` }), { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }) }
        const td = await resp.json()
        const mr = await fetch(`https://graph.microsoft.com/v1.0/users/${mailbox}/messages?$top=1`, { headers: { Authorization: `Bearer ${td.access_token}` } })
        if (!mr.ok) return new Response(JSON.stringify({ success: false, error: `Buzon ${mailbox} no accesible.` }), { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
        return new Response(JSON.stringify({ success: true, message: `Conexion Graph API exitosa.` }), { headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
      }

      if (provider === 'smtp') {
        const h = (body.smtp_host as string) || (settings.smtp_host as string) || Deno.env.get('SMTP_HOST')
        const u = (body.smtp_user as string) || (settings.smtp_user as string) || Deno.env.get('SMTP_USER')
        const p = (body.smtp_password as string) || Deno.env.get('SMTP_PASSWORD')
        const port = Number(body.smtp_port) || Number(settings.smtp_port) || 587
        if (!h || !u || !p) {
          const f = []; if (!h) f.push('Servidor SMTP'); if (!u) f.push('Usuario SMTP'); if (!p) f.push('Contrasena SMTP')
          return new Response(JSON.stringify({ success: false, error: `Credenciales SMTP incompletas: ${f.join(', ')}.` }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
        }
        try {
          const conn = await Deno.connectTls({ hostname: h, port })
          const dec = new TextDecoder(); const enc = new TextEncoder()
          async function rd(): Promise<string> { const b = new Uint8Array(4096); const n = await conn.read(b); if (n === null) throw new Error('Conexion cerrada'); return dec.decode(b.subarray(0, n)) }
          async function sc(cmd: string, ec: number): Promise<string> {
            if (cmd) await conn.write(enc.encode(cmd + '\r\n')); const r = await rd(); const c = parseInt(r.substring(0, 3))
            if (c !== ec) { const cn = cmd.split(' ')[0]; if (cn === 'AUTH' && c === 535) throw new Error('Error de autenticacion SMTP: credenciales rechazadas.'); throw new Error(`Error SMTP (${cn || 'conexion'}): esperado ${ec}, recibido ${c}`) }
            return r
          }
          try {
            await sc('', 220); await sc(`EHLO ${h}`, 250); await sc('STARTTLS', 220)
            const tc = await Deno.startTls(conn, { hostname: h })
            const trd = async () => { const b = new Uint8Array(4096); const n = await tc.read(b); if (n === null) throw new Error('TLS cerrado'); return dec.decode(b.subarray(0, n)) }
            const tsc = async (cmd: string, ec: number): Promise<string> => { if (cmd) await tc.write(enc.encode(cmd + '\r\n')); const r = await trd(); const c = parseInt(r.substring(0, 3)); if (c !== ec) throw new Error(`SMTP TLS: esperado ${ec}, recibido ${c}`); return r }
            await tsc(`EHLO ${h}`, 250); await tsc('AUTH LOGIN', 334); await tsc(btoa(u), 334); await tsc(btoa(p), 235); await tsc('QUIT', 221); tc.close()
            return new Response(JSON.stringify({ success: true, message: `Conexion SMTP exitosa con ${h}:${port}.` }), { headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
          } finally { try { conn.close() } catch { } }
        } catch (err) { return new Response(JSON.stringify({ success: false, error: err instanceof Error ? err.message : 'Error SMTP' }), { status: 502, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }) }
      }
      return new Response(JSON.stringify({ error: 'Proveedor no soportado.' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
    }

    if (action === 'update-config') {
      const nc = body.config as Record<string, unknown> | undefined
      if (!nc) return new Response(JSON.stringify({ error: 'Falta el objeto config' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
      const allowed = ['provider', 'mailbox', 'from_name', 'smtp_host', 'smtp_port', 'smtp_user', 'imap_host', 'imap_port', 'imap_user']
      const ts: Record<string, unknown> = {}; for (const k of allowed) { if (k in nc) ts[k] = nc[k] }
      if (ts.provider && !['graph_api', 'smtp', 'none'].includes(ts.provider as string)) return new Response(JSON.stringify({ error: 'Proveedor invalido' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
      if (ts.smtp_port != null) { const p = Number(ts.smtp_port); if (isNaN(p) || p < 1 || p > 65535) return new Response(JSON.stringify({ error: 'Puerto SMTP invalido' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }); ts.smtp_port = p }
      if (ts.imap_port != null) { const p = Number(ts.imap_port); if (isNaN(p) || p < 1 || p > 65535) return new Response(JSON.stringify({ error: 'Puerto IMAP invalido' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }); ts.imap_port = p }
      ts.updated_at = new Date().toISOString()
      const { error: dbe } = await supabase.from('email_settings').upsert({ id: 1, ...ts }, { onConflict: 'id' })
      if (dbe) return new Response(JSON.stringify({ error: `Error BD: ${dbe.message}` }), { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
      const st = await readSettings(); return new Response(JSON.stringify({ success: true, message: 'Configuracion guardada', saved: ts, status: buildStatus(st) }), { headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
    }
    return new Response(JSON.stringify({ error: 'Accion no soportada' }), { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
  } catch (error) {
    console.error('email-config error:', error)
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : 'Error interno' }), { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
  }
})
