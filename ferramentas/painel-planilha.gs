/**
 * Avatar Milionário - liberação de acesso pela Kiwify (grátis, roda no Google).
 *
 * Como instalar (uma vez só):
 * 1. Crie uma planilha nova no Google Planilhas.
 * 2. Menu Extensões > Apps Script. Apague o que estiver lá e cole este arquivo inteiro.
 * 3. Preencha SENHA_PAINEL e PRODUTO_BUMP aqui embaixo.
 * 4. Implantar > Nova implantação > tipo "App da Web".
 *    Executar como: Eu. Quem pode acessar: Qualquer pessoa. Implantar e autorizar.
 * 5. Copie a URL do app da web (termina em /exec). Ela é usada na Kiwify, na área de membros e no painel.
 *
 * A Kiwify chama este script a cada evento de compra (POST).
 * A área de membros pergunta se um e-mail tem acesso (GET ?email=...).
 * O seu painel lista os alunos e libera ou bloqueia na mão (POST com a senha).
 */

/* Token do webhook que a Kiwify mostra ao criar o webhook. Vazio = não confere a assinatura. */
const TOKEN_KIWIFY = '';
/* Senha do seu painel de alunos. Vazio = painel desligado. */
const SENHA_PAINEL = '';
/* Parte do nome (ou o ID) do produto do order bump na Kiwify, ex.: 'TikTok Shop'. Vazio = sem bump. */
const PRODUTO_BUMP = 'Virais';

const ABA = 'Alunos';
const CABECALHO = ['email', 'nome', 'acesso', 'bump', 'evento', 'pedido', 'atualizado em', 'observacao'];
const COLS = CABECALHO.length;

/* Status e eventos da Kiwify que liberam e que bloqueiam */
const LIBERA = ['paid', 'approved', 'order_approved', 'compra_aprovada'];
const BLOQUEIA = ['refunded', 'chargedback', 'chargeback', 'order_refunded', 'compra_reembolsada', 'subscription_canceled', 'canceled', 'cancelled', 'order_canceled'];
const PENDENTE = ['waiting_payment', 'pending', 'refused', 'pix_created', 'billet_created', 'pix_gerado', 'boleto_gerado', 'compra_recusada'];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(CABECALHO);
    sh.setFrozenRows(1);
  } else if (!sh.getRange(1, COLS).getValue()) {
    /* Planilha antiga, de antes da coluna de observação */
    sh.getRange(1, COLS).setValue(CABECALHO[COLS - 1]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Linha (1-based) do e-mail na planilha, ou 0 se não existir */
function acharLinha_(sh, email) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const emails = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (String(emails[i][0]).trim().toLowerCase() === email) return i + 2;
  }
  return 0;
}

function assinaturaOk_(corpo, assinatura) {
  if (!TOKEN_KIWIFY) return true;
  const bytes = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, corpo, TOKEN_KIWIFY);
  const hex = bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  return hex === String(assinatura || '').toLowerCase();
}

/* Grava ou atualiza a linha do e-mail. Campos ausentes (undefined) mantêm o valor que já estava; null apaga. */
function salvar_(email, campos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = aba_();
    const r = acharLinha_(sh, email);
    const atual = r ? sh.getRange(r, 1, 1, COLS).getValues()[0] : [email, '', '', '', '', '', '', ''];
    const pega = (v, i) => (v === null ? '' : v === undefined || v === '' ? atual[i] : v);
    const linha = [
      email,
      pega(campos.nome, 1),
      pega(campos.acesso, 2),
      pega(campos.bump, 3),
      pega(campos.evento, 4),
      pega(campos.pedido, 5),
      new Date(),
      pega(campos.observacao, 7)
    ];
    if (r) sh.getRange(r, 1, 1, COLS).setValues([linha]);
    else sh.appendRow(linha);
  } finally {
    lock.releaseLock();
  }
}

function listar_() {
  const sh = aba_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  const sit = situacaoEmails_();
  const ips = ipsPorEmail_();
  return sh.getRange(2, 1, n - 1, COLS).getValues().map(v => {
    const s = sit[String(v[0]).trim().toLowerCase()];
    const ip = ips[String(v[0]).trim().toLowerCase()] || [];
    return {
      email: v[0], nome: v[1], acesso: v[2], bump: v[3], evento: v[4], pedido: v[5],
      observacao: v[7] || '', ips: ip.length, listaIps: ip,
      quando: v[6] instanceof Date ? v[6].toISOString() : String(v[6]),
      mail: s ? { quando: new Date(s.quando).toISOString(), ok: s.ok, erro: s.erro, entregue: s.entregue, aberto: s.aberto ? new Date(s.aberto).toISOString() : '', problema: s.problema } : null
    };
  }).reverse();
}

/* Painel: { acao, senha, email?, nome? }
   acao: listar | liberar | bloquear | bloquear_ips | liberar_bump | bloquear_bump | ips */
function painel_(d) {
  if (!SENHA_PAINEL || d.senha !== SENHA_PAINEL) {
    Utilities.sleep(1500);
    return json_({ ok: false, erro: 'senha' });
  }
  const email = String(d.email || '').trim().toLowerCase();
  const mudancas = {
    liberar: { acesso: 'ativo' },
    bloquear: { acesso: 'bloqueado' },
    /* Bloqueio por IPs: o aluno vê o aviso para chamar o suporte no WhatsApp */
    bloquear_ips: { acesso: 'bloqueado', observacao: OBS_IPS + ' em ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') },
    liberar_bump: { bump: 'ativo' },
    bloquear_bump: { bump: 'bloqueado' }
  }[d.acao];
  if (mudancas) {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, erro: 'e-mail inválido' });
    mudancas.nome = String(d.nome || '').trim();
    mudancas.evento = 'manual';
    /* Liberar na mão tira o bloqueio de IPs: apaga a observação e zera os IPs da conta */
    if (d.acao === 'liberar') { mudancas.observacao = null; limparIps_(email); }
    /* Liberar com a caixinha do Vídeos Virais marcada: libera os dois de uma vez (um e-mail só) */
    if (d.acao === 'liberar' && d.comBump) mudancas.bump = 'ativo';
    const antes = estado_(email);
    salvar_(email, mudancas);
    /* Liberou pelo painel: a pessoa recebe o mesmo e-mail de boas-vindas da compra (só na primeira vez) */
    let envio = null;
    try {
      /* Desbloqueio de quem foi barrado por IPs não reenvia o e-mail de boas-vindas */
      if (d.acao === 'liberar' && antes.acesso !== 'ativo' && antes.observacao.indexOf(OBS_IPS) !== 0) envio = enviarBoasVindas_(email, mudancas.nome || antes.nome, true, mudancas.bump === 'ativo' || antes.bump === 'ativo');
      if (d.acao === 'liberar_bump' && antes.bump !== 'ativo') envio = enviarBoasVindas_(email, mudancas.nome || antes.nome, false, true);
    } catch (x) { envio = { ok: false, erro: String(x.message || x) }; }
    if (d.acao !== 'bloquear' && d.acao !== 'bloquear_bump' && d.acao !== 'bloquear_ips') return json_({ ok: true, envio: envio, bump: !!PRODUTO_BUMP, alunos: listar_() });
  }
  /* Botão "Reenviar e-mail" do painel: manda de novo o e-mail de boas-vindas de verdade */
  if (d.acao === 'reenviar') {
    const est = estado_(email);
    if (est.acesso !== 'ativo' && est.bump !== 'ativo') return json_({ ok: false, erro: 'Esse e-mail não tem acesso liberado.' });
    const envio = enviarBoasVindas_(email, est.nome, est.acesso === 'ativo', est.bump === 'ativo');
    return json_({ ok: true, envio: envio, bump: !!PRODUTO_BUMP, alunos: listar_() });
  }
  if (d.acao === 'visitas') return json_({ ok: true, stats: estatisticas_() });
  if (d.acao === 'aovivo') return json_({ ok: true, aoVivo: aoVivo_() });
  if (d.acao === 'ips') return json_({ ok: true, ips: resumoIps_() });
  return json_({ ok: true, bump: !!PRODUTO_BUMP, alunos: listar_() });
}

/* Junta nome e ID de todos os produtos que aparecem no pedido (produto principal e bumps) */
function produtos_(d) {
  const achados = [];
  const visitar = (o, prof) => {
    if (!o || typeof o !== 'object' || prof > 4) return;
    if (Array.isArray(o)) { o.forEach(x => visitar(x, prof + 1)); return; }
    const nome = o.product_name || o.name;
    const id = o.product_id || o.id;
    if (o.product_name || o.product_id) achados.push(String(nome || '') + ' ' + String(id || ''));
    Object.keys(o).forEach(k => {
      if (/product|bump|item/i.test(k)) visitar(o[k], prof + 1);
    });
  };
  visitar(d, 0);
  return achados;
}

/* ---------- VISITAS ---------- */
const ABA_EV = 'Eventos';
const CAB_EV = ['data', 'tipo', 'pagina', 'afiliado', 'origem', 'visitante', 'visita', 'segundos', 'cliques'];
const TIPOS = ['visita', 'checkout', 'fim'];

function abaEv_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_EV);
  if (!sh) {
    sh = ss.insertSheet(ABA_EV);
    sh.appendRow(CAB_EV);
    sh.setFrozenRows(1);
  }
  return sh;
}

const corta_ = (v, n) => String(v || '').slice(0, n);

/* ---------- E-MAIL DE BOAS-VINDAS ----------
   Enviado quando a compra é aprovada. Com a chave da Brevo preenchida, sai pela Brevo
   como contato@avatarmilionario.com (grátis até 300/dia). Sem a chave, sai pelo seu Gmail. */
const BREVO_KEY = '';
const REMETENTE = { name: 'Avatar Milionário', email: 'contato@avatarmilionario.com' };
const RESPONDER_PARA = 'alexandrej7888@gmail.com';
const AREA_MEMBROS = 'https://minhapagina.pro/area-de-membros/';
const WHATS_SUPORTE = '5538999715362';
const IMG_EMAIL = 'https://cdn.jsdelivr.net/gh/alexandrej7888-hue/avatar-milionario-imagens@c8bb30b1cdc0463cc9c411e0ede26f1fc676f2b9/capas/email-boasvindas.jpg';
function estado_(email) {
  const sh = aba_();
  const r = acharLinha_(sh, email);
  if (!r) return { acesso: '', bump: '', nome: '', observacao: '' };
  const v = sh.getRange(r, 1, 1, COLS).getValues()[0];
  return { acesso: String(v[2]).toLowerCase(), bump: String(v[3]).toLowerCase(), nome: String(v[1] || '').trim(), observacao: String(v[7] || '') };
}
function passo_(n, titulo, texto) {
  return '<tr><td style="padding:0 0 14px"><table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr>' +
    '<td width="38" valign="top"><div style="width:30px;height:30px;line-height:30px;border-radius:50%;background:#22c55e;color:#03150a;font-weight:bold;text-align:center;font-size:14px">' + n + '</div></td>' +
    '<td style="color:#cfe3d6;font-size:15px;line-height:1.5"><b style="color:#fff">' + titulo + '</b><br>' + texto + '</td></tr></table></td></tr>';
}
function emailHtml_(email, titulo, intro, passos, extra) {
  return '<div style="background:#050907;padding:24px 10px;font-family:Arial,Helvetica,sans-serif">' +
    '<table role="presentation" align="center" width="100%" cellspacing="0" cellpadding="0" style="max-width:560px;background:#0d1611;border:1px solid #24392c;border-radius:18px;overflow:hidden">' +
    '<tr><td><img src="' + IMG_EMAIL + '" width="560" alt="Avatar Milionário" style="display:block;width:100%;height:auto;border:0"></td></tr>' +
    '<tr><td style="padding:28px 26px 8px;color:#ecf5ef">' +
      '<div style="font-size:12px;letter-spacing:3px;color:#4ade80;font-weight:bold">AVATAR MILIONÁRIO</div>' +
      '<h1 style="font-size:24px;line-height:1.25;margin:10px 0 12px;color:#fff">' + titulo + '</h1>' +
      '<p style="font-size:15px;line-height:1.6;color:#cfe3d6;margin:0 0 18px">' + intro + '</p>' +
      (passos ? '<table role="presentation" width="100%" cellspacing="0" cellpadding="0">' + passos + '</table>' : '') +
    '</td></tr>' +
    '<tr><td align="center" style="padding:8px 26px 6px"><a href="' + AREA_MEMBROS + '" style="background:#22c55e;color:#03150a;text-decoration:none;font-weight:bold;font-size:16px;padding:16px 30px;border-radius:12px;display:inline-block">Entrar na área de membros</a></td></tr>' +
    '<tr><td style="padding:12px 26px 6px;text-align:center;color:#93a89b;font-size:13px;line-height:1.5">Seu login é este e-mail: <b style="color:#ecf5ef">' + email + '</b><br>Não precisa de senha.</td></tr>' +
    (extra || '') +
    '<tr><td style="padding:18px 26px 26px;text-align:center;color:#62756a;font-size:12px;line-height:1.5">' +
      (WHATS_SUPORTE ? '<a href="https://wa.me/' + WHATS_SUPORTE + '" style="background:#25d366;color:#03150a;text-decoration:none;font-weight:bold;font-size:14px;padding:11px 20px;border-radius:10px;display:inline-block;margin-bottom:12px">Suporte no WhatsApp</a><br>Dúvidas? Chame no WhatsApp (38) 99971-5362 ou responda este e-mail.<br>' : 'Dúvidas? É só responder este e-mail.<br>') +
      'Você recebeu este e-mail porque comprou o Avatar Milionário.</td></tr>' +
    '</table></div>';
}
/* Manda pela Brevo (se tiver chave) ou pelo Gmail. Devolve { ok, via, erro } e anota na aba E-mails. */
function mandar_(para, nome, assunto, html) {
  let res = null;
  if (BREVO_KEY) {
    try {
      const r = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'api-key': BREVO_KEY, accept: 'application/json' },
        payload: JSON.stringify({ sender: REMETENTE, to: [{ email: para, name: nome || para }], replyTo: { email: RESPONDER_PARA }, subject: assunto, htmlContent: html })
      });
      if (r.getResponseCode() < 300) res = { ok: true, via: 'brevo' };
      else console.error('Brevo recusou (' + r.getResponseCode() + '): ' + r.getContentText());
    } catch (x) { console.error('Brevo: ' + x); }
  }
  if (!res) {
    try {
      MailApp.sendEmail({ to: para, name: REMETENTE.name, replyTo: RESPONDER_PARA, subject: assunto, htmlBody: html });
      res = { ok: true, via: 'gmail' };
    } catch (x) { res = { ok: false, via: '', erro: String(x.message || x) }; }
  }
  anotarEmail_(para, assunto, res);
  return res;
}

/* ---------- HISTÓRICO DE E-MAILS ---------- */
const ABA_EMAILS = 'E-mails';
const CAB_EMAILS = ['data', 'email', 'assunto', 'enviado', 'via', 'erro'];
function abaEmails_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_EMAILS);
  if (!sh) { sh = ss.insertSheet(ABA_EMAILS); sh.appendRow(CAB_EMAILS); sh.setFrozenRows(1); }
  return sh;
}
function anotarEmail_(para, assunto, res) {
  try { abaEmails_().appendRow([new Date(), String(para).toLowerCase(), assunto, res.ok ? 'sim' : 'não', res.via || '', res.erro || '']); } catch (x) { console.error('anotar e-mail: ' + x); }
}
/* Situação do último e-mail de cada aluno: enviado/erro (planilha) + entregue/aberto (Brevo) */
function situacaoEmails_() {
  const sh = abaEmails_();
  const n = sh.getLastRow();
  const sit = {};
  (n < 2 ? [] : sh.getRange(2, 1, n - 1, CAB_EMAILS.length).getValues()).forEach(r => {
    const em = String(r[1]).toLowerCase();
    const t = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
    sit[em] = { quando: t, ok: r[3] === 'sim', via: r[4], erro: r[5], entregue: false, aberto: 0, problema: '' };
  });
  if (BREVO_KEY && Object.keys(sit).length) {
    try {
      const r = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/statistics/events?limit=2500&days=90&sort=desc', {
        muteHttpExceptions: true, headers: { 'api-key': BREVO_KEY, accept: 'application/json' }
      });
      if (r.getResponseCode() < 300) {
        (JSON.parse(r.getContentText()).events || []).forEach(ev => {
          const s = sit[String(ev.email || '').toLowerCase()];
          if (!s) return;
          const t = new Date(ev.date).getTime();
          if (t < s.quando - 60000) return; /* só eventos do último envio */
          const tipo = String(ev.event || '');
          if (tipo === 'delivered') s.entregue = true;
          if (tipo === 'opened' || tipo === 'uniqueOpened' || tipo === 'loadedByProxy' || tipo === 'clicks') { s.entregue = true; s.aberto = Math.max(s.aberto, t); }
          if (/bounce|blocked|invalid|spam|error/i.test(tipo)) s.problema = tipo;
        });
      }
    } catch (x) { console.error('eventos Brevo: ' + x); }
  }
  return sit;
}
/* Guarda o comprador como contato na Brevo (para campanhas depois) */
function contatoBrevo_(email, nome, bump) {
  if (!BREVO_KEY) return;
  try {
    UrlFetchApp.fetch('https://api.brevo.com/v3/contacts', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'api-key': BREVO_KEY, accept: 'application/json' },
      payload: JSON.stringify({ email: email, updateEnabled: true, attributes: { FIRSTNAME: String(nome || '').split(/\s+/)[0] || '', COMPROU_BUMP: !!bump } })
    });
  } catch (x) { console.error('contato Brevo: ' + x); }
}
function enviarBoasVindas_(email, nome, curso, bump) {
  const primeiro = String(nome || '').trim().split(/\s+/)[0] || '';
  const oi = primeiro ? 'Oi, ' + primeiro + '!' : 'Oi!';
  let res = { ok: false, erro: 'nada para enviar' };
  try {
    if (curso) {
      const passos =
        passo_(1, 'Entre na área de membros', 'Clique no botão abaixo e digite o mesmo e-mail que você usou na compra.') +
        passo_(2, 'Comece por "Comece por aqui"', 'Assista as aulas na ordem. São curtas e você já faz junto.') +
        passo_(3, 'Monte seu avatar', 'Use o Criador de avatar e os prompts prontos para criar seu personagem.') +
        passo_(4, 'Poste seu primeiro vídeo', (bump ? 'Pegue um vídeo do seu <b style="color:#facc15">pack de Vídeos Virais</b> e coloque seu avatar nele.' : 'Escolha um vídeo viral de referência e coloque seu avatar nele.'));
      const extra = bump ? '' : '<tr><td style="padding:14px 26px 0"><div style="border:1px dashed #facc15;border-radius:12px;padding:14px;color:#cfe3d6;font-size:14px;line-height:1.5;text-align:center">🔥 Quer vídeos prontos para postar? O <b style="color:#facc15">pack Vídeos Virais</b> está disponível dentro da área de membros.</div></td></tr>';
      res = mandar_(email, nome, primeiro ? primeiro + ', seu acesso ao Avatar Milionário chegou 🚀' : 'Seu acesso ao Avatar Milionário chegou 🚀',
        emailHtml_(email, 'Bem-vindo ao Avatar Milionário! 🎉', oi + ' Sua compra foi aprovada e o seu acesso já está liberado. Veja como começar:', passos, extra));
    } else if (bump) {
      res = mandar_(email, nome, 'Seu pack de Vídeos Virais foi liberado 🔥',
        emailHtml_(email, 'Pack de Vídeos Virais liberado! 🔥', oi + ' O pack já está na sua área de membros, no menu <b>Vídeos virais</b>. É só baixar, colocar seu avatar e postar.', '', ''));
    }
    contatoBrevo_(email, nome, bump);
  } catch (x) { console.error('e-mail não enviado: ' + x); res = { ok: false, erro: String(x.message || x) }; }
  return res;
}

/* Rode esta função uma vez no editor (botão Executar) para autorizar o envio de e-mail.
   Ela manda um e-mail de exemplo para a sua própria conta. */
function testarEmail() {
  /* Chamadas diretas para o Google pedir todas as permissões de uma vez */
  UrlFetchApp.fetch('https://api.brevo.com/v3/account', { muteHttpExceptions: true, headers: { 'api-key': BREVO_KEY } });
  MailApp.getRemainingDailyQuota();
  enviarBoasVindas_(Session.getEffectiveUser().getEmail(), 'Teste', true, true);
}

/* ---------- AO VIVO ----------
   Cada página aberta manda um "ping" a cada 45s. Guardamos no cache (sem escrever na planilha)
   quem mandou ping no último 1min30: isso é quem está na página agora. */
const VIVO_SEG = 90;
function vivoLer_() {
  const c = CacheService.getScriptCache().get('vivo');
  try { return c ? JSON.parse(c) : {}; } catch (x) { return {}; }
}
function vivoPing_(d) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    const agora = Date.now(), m = vivoLer_(), novo = {};
    Object.keys(m).forEach(k => { if (agora - m[k][0] < VIVO_SEG * 1000) novo[k] = m[k]; });
    novo[corta_(d.visita, 24) || Math.random()] = [agora, corta_(d.afiliado, 40).toLowerCase(), corta_(d.pagina, 60)];
    CacheService.getScriptCache().put('vivo', JSON.stringify(novo), 600);
  } finally { lock.releaseLock(); }
}
function aoVivo_() {
  const agora = Date.now(), m = vivoLer_(), por = {};
  let total = 0;
  Object.keys(m).forEach(k => {
    if (agora - m[k][0] >= VIVO_SEG * 1000) return;
    total++;
    const a = m[k][1] || '(sem afiliado)';
    por[a] = (por[a] || 0) + 1;
  });
  return { total: total, afiliados: Object.keys(por).map(k => ({ nome: k, pessoas: por[k] })).sort((a, b) => b.pessoas - a.pessoas) };
}

/* ---------- VENDAS ---------- */
const ABA_VENDAS = 'Vendas';
const CAB_VENDAS = ['data', 'status', 'produto', 'valor', 'sua comissao', 'afiliado', 'pagina', 'origem', 'campanha', 'email', 'pedido'];
function abaVendas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_VENDAS);
  if (!sh) { sh = ss.insertSheet(ABA_VENDAS); sh.appendRow(CAB_VENDAS); sh.setFrozenRows(1); }
  return sh;
}
/* A Kiwify manda valores em centavos (ex.: 4700 = R$ 47,00) */
const reais_ = v => { const n = Number(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? (Number.isInteger(n) ? n / 100 : n) : 0; };
function registrarVenda_(d, novo, email) {
  const com = d.Commissions || d.commissions || {};
  const tr = d.TrackingParameters || d.tracking_parameters || {};
  const prod = (d.Product && (d.Product.product_name || d.Product.name)) || '';
  /* Afiliado que a Kiwify comissionou (se houver) */
  const lojas = com.commissioned_stores || [];
  const afi = lojas.filter(s => /affiliate|afiliad/i.test(s.type || '')).map(s => s.custom_name || s.email || s.id)[0] || '';
  const src = String(tr.src || '').toLowerCase();
  const linha = [
    new Date(), novo === 'ativo' ? 'aprovada' : 'reembolsada', prod,
    reais_(com.charge_amount != null ? com.charge_amount : d.charge_amount),
    reais_(com.my_commission != null ? com.my_commission : ''),
    src && src !== 'principal' ? src : (afi ? String(afi).toLowerCase() : ''),
    src, String(tr.utm_source || '').toLowerCase(), String(tr.utm_campaign || ''), email,
    d.order_id || d.order_ref || ''
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    /* A Kiwify reenvia o mesmo aviso várias vezes: o mesmo pedido com o mesmo status só entra uma vez */
    const sh = abaVendas_();
    const n = sh.getLastRow();
    const ja = n < 2 ? [] : sh.getRange(2, 1, n - 1, CAB_VENDAS.length).getValues();
    if (linha[10] && ja.some(r => String(r[10]) === String(linha[10]) && r[1] === linha[1])) return;
    sh.appendRow(linha);
  } finally { lock.releaseLock(); }
}

/* Evento enviado pelas páginas: { ev, pagina, afiliado, origem, visitante, visita, segundos, cliques } */
function registrar_(d) {
  if (d.ev === 'ping') { vivoPing_(d); return json_({ ok: true }); }
  if (TIPOS.indexOf(d.ev) < 0) return json_({ ok: false });
  const linha = [
    new Date(), d.ev, corta_(d.pagina, 80), corta_(d.afiliado, 40).toLowerCase(), corta_(d.origem, 60).toLowerCase(),
    corta_(d.visitante, 24), corta_(d.visita, 24),
    Math.max(0, Math.min(3600, Math.round(+d.segundos || 0))), Math.max(0, Math.min(500, Math.round(+d.cliques || 0)))
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { abaEv_().appendRow(linha); } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

/* Soma os eventos por período, por afiliado, por origem e por dia */
function estatisticas_() {
  const sh = abaEv_();
  const n = sh.getLastRow();
  const rows = n < 2 ? [] : sh.getRange(2, 1, n - 1, CAB_EV.length).getValues();
  const agora = new Date();
  const hoje0 = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate()).getTime();
  const DIA = 86400000;
  const periodos = { hoje: hoje0, '7d': hoje0 - 6 * DIA, '30d': hoje0 - 29 * DIA, total: 0 };

  const novo = () => ({ visitas: 0, visitantes: {}, cliques: 0, segundos: 0, fins: 0, checkouts: 0, visitasCheckout: {} });
  const fecha = a => {
    const v = Object.keys(a.visitantes).length;
    return {
      visitas: a.visitas, visitantes: v, cliques: a.cliques, checkouts: a.checkouts,
      pessoasCheckout: Object.keys(a.visitasCheckout).length,
      tempoMedio: a.fins ? Math.round(a.segundos / a.fins) : 0,
      conversao: a.visitas ? Math.round(Object.keys(a.visitasCheckout).length / a.visitas * 1000) / 10 : 0
    };
  };
  const soma = (a, r) => {
    const tipo = r[1];
    if (tipo === 'visita') { a.visitas++; if (r[5]) a.visitantes[r[5]] = 1; }
    if (tipo === 'checkout') { a.checkouts++; a.visitasCheckout[r[6] || Math.random()] = 1; }
    if (tipo === 'fim') { a.fins++; a.segundos += +r[7] || 0; a.cliques += +r[8] || 0; }
  };

  const saida = {};
  Object.keys(periodos).forEach(p => {
    const ini = periodos[p];
    const tot = novo(), afi = {}, ori = {}, pag = {};
    rows.forEach(r => {
      const t = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
      if (t < ini) return;
      soma(tot, r);
      const a = r[3] || '(sem afiliado)';
      soma(afi[a] = afi[a] || novo(), r);
      const o = r[4] || '(direto)';
      soma(ori[o] = ori[o] || novo(), r);
      const g = r[2] || '/';
      soma(pag[g] = pag[g] || novo(), r);
    });
    const lista = obj => Object.keys(obj).map(k => Object.assign({ nome: k }, fecha(obj[k]))).sort((x, y) => y.visitas - x.visitas);
    saida[p] = { total: fecha(tot), afiliados: lista(afi), origens: lista(ori), paginas: lista(pag) };
  });

  /* Visitas e checkouts por dia nos últimos 30 dias */
  const dias = [];
  for (let i = 29; i >= 0; i--) dias.push({ dia: Utilities.formatDate(new Date(hoje0 - i * DIA), Session.getScriptTimeZone(), 'dd/MM'), ini: hoje0 - i * DIA, visitas: 0, checkouts: 0 });
  rows.forEach(r => {
    const t = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
    const i = Math.floor((t - dias[0].ini) / DIA);
    if (i < 0 || i >= dias.length) return;
    if (r[1] === 'visita') dias[i].visitas++;
    if (r[1] === 'checkout') dias[i].checkouts++;
  });
  saida.dias = dias.map(d => ({ dia: d.dia, visitas: d.visitas, checkouts: d.checkouts }));

  /* Vendas por período, afiliado e origem, com conversão sobre as visitas */
  const shV = abaVendas_();
  const nv = shV.getLastRow();
  const vistos = {};
  const vendas = (nv < 2 ? [] : shV.getRange(2, 1, nv - 1, CAB_VENDAS.length).getValues()).filter(r => {
    /* Ignora avisos repetidos do mesmo pedido (linhas antigas gravadas antes da correção) */
    if (!r[10]) return true;
    const k = r[10] + '|' + r[1];
    if (vistos[k]) return false;
    return (vistos[k] = true);
  });
  const tempo = x => x instanceof Date ? x.getTime() : new Date(x).getTime();
  Object.keys(periodos).forEach(p => {
    const ini = periodos[p];
    const soma = { vendas: 0, valor: 0, comissao: 0, reembolsos: 0 };
    const afi = {}, ori = {};
    const conta = (obj, k, r) => {
      const a = obj[k] = obj[k] || { vendas: 0, valor: 0, comissao: 0, reembolsos: 0 };
      if (r[1] === 'aprovada') { a.vendas++; a.valor += +r[3] || 0; a.comissao += +r[4] || 0; } else a.reembolsos++;
    };
    vendas.forEach(r => {
      if (tempo(r[0]) < ini) return;
      if (r[1] === 'aprovada') { soma.vendas++; soma.valor += +r[3] || 0; soma.comissao += +r[4] || 0; } else soma.reembolsos++;
      conta(afi, r[5] || '(sem afiliado)', r);
      conta(ori, r[7] || '(direto)', r);
    });
    const vis = saida[p].total.visitas;
    soma.conversao = vis ? Math.round(soma.vendas / vis * 1000) / 10 : 0;
    const visAfi = {}; saida[p].afiliados.forEach(a => { visAfi[a.nome] = a.visitas; });
    const lista = (obj, comVis) => Object.keys(obj).map(k => {
      const o = Object.assign({ nome: k }, obj[k]);
      if (comVis) { o.visitas = visAfi[k] || 0; o.conversao = o.visitas ? Math.round(o.vendas / o.visitas * 1000) / 10 : 0; }
      return o;
    }).sort((a, b) => b.valor - a.valor);
    saida[p].vendas = { total: soma, afiliados: lista(afi, true), origens: lista(ori, false) };
  });
  saida.ultimasVendas = vendas.slice(-20).reverse().map(r => ({
    quando: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), status: r[1], produto: r[2], valor: r[3],
    afiliado: r[5], origem: r[7], email: r[9]
  }));
  saida.aoVivo = aoVivo_();
  return saida;
}

/* Kiwify, páginas ou painel -> planilha */
function doPost(e) {
  const corpo = (e && e.postData && e.postData.contents) || '';
  let d;
  try { d = JSON.parse(corpo); } catch (x) { return json_({ ok: false, erro: 'json' }); }
  if (d.ev) return registrar_(d);
  if (d.acao) return painel_(d);

  if (!assinaturaOk_(corpo, e.parameter && e.parameter.signature)) return json_({ ok: false, erro: 'assinatura' });
  if (d.order && typeof d.order === 'object') d = d.order;

  const cli = d.Customer || d.customer || {};
  const email = String(cli.email || d.email || '').trim().toLowerCase();
  if (!email) return json_({ ok: false, erro: 'sem e-mail' });

  const status = String(d.order_status || d.status || '').toLowerCase();
  const evento = String(d.webhook_event_type || d.event || '').toLowerCase();
  let novo = '';
  if (LIBERA.indexOf(status) >= 0 || LIBERA.indexOf(evento) >= 0) novo = 'ativo';
  if (BLOQUEIA.indexOf(status) >= 0 || BLOQUEIA.indexOf(evento) >= 0) novo = 'bloqueado';
  /* Pix/boleto gerado ou cartão recusado não é venda: só libera com pagamento confirmado */
  if (novo === 'ativo' && PENDENTE.indexOf(status) >= 0) novo = '';
  if (!novo) return json_({ ok: true, ignorado: status || evento });

  /* Separa o que é bump do que é o curso */
  const prods = produtos_(d);
  const ehBump = p => PRODUTO_BUMP && p.toLowerCase().indexOf(PRODUTO_BUMP.toLowerCase()) >= 0;
  const temBump = prods.some(ehBump);
  const temCurso = !prods.length || prods.some(p => !ehBump(p));

  const nome = cli.full_name || cli.name || [cli.first_name, cli.last_name].filter(Boolean).join(' ') || '';
  const campos = { nome: nome, evento: evento || status, pedido: d.order_id || d.order_ref || d.id || '' };
  if (temCurso) campos.acesso = novo;
  if (temBump) campos.bump = novo;
  /* Reembolso/chargeback/cancelamento do curso corta tudo (área de membros e Vídeos Virais) */
  if (novo === 'bloqueado' && temCurso) campos.bump = 'bloqueado';
  const antes = estado_(email);
  salvar_(email, campos);
  /* Só manda e-mail na primeira liberação (não repete se a Kiwify reenviar o aviso) */
  if (novo === 'ativo') {
    const cursoNovo = temCurso && antes.acesso !== 'ativo';
    const bumpNovo = temBump && antes.bump !== 'ativo';
    if (cursoNovo || bumpNovo) enviarBoasVindas_(email, nome, cursoNovo, bumpNovo || antes.bump === 'ativo');
  }
  registrarVenda_(d, novo, email);
  return json_({ ok: true, email: email, acesso: campos.acesso || null, bump: campos.bump || null });
}

/* ---------- IPs ----------
   A área de membros manda o IP de quem entra. Cada e-mail pode usar até LIMITE_IPS IPs diferentes;
   passou disso, a conta ganha a observação de alerta e aparece no painel com as cidades de cada IP.
   Quem bloqueia é você, pelo painel (botão "Bloquear por IPs"): aí a pessoa vê o aviso para chamar o suporte.
   Liberar pelo painel zera os IPs e tira a observação. */
const LIMITE_IPS = 2;
const ABA_IPS = 'IPs';
const CAB_IPS = ['email', 'ip', 'primeiro acesso', 'ultimo acesso', 'acessos', 'local'];
const OBS_ALERTA = 'ALERTA: acesso de vários IPs';
const OBS_IPS = 'Bloqueado por acesso de vários IPs';
/* Cidade/estado do IP (serviços grátis, só é chamado na primeira vez que o IP aparece) */
function localIp_(ip) {
  try {
    const r = UrlFetchApp.fetch('https://ipwho.is/' + ip + '?lang=pt-BR&fields=success,city,region,country', { muteHttpExceptions: true });
    const j = JSON.parse(r.getContentText());
    if (j.success) return [j.city, j.region, j.country].filter(Boolean).join(', ');
  } catch (x) { console.error('ipwho.is: ' + x); }
  try {
    const r = UrlFetchApp.fetch('http://ip-api.com/json/' + ip + '?lang=pt-BR&fields=status,city,regionName,country', { muteHttpExceptions: true });
    const j = JSON.parse(r.getContentText());
    if (j.status === 'success') return [j.city, j.regionName, j.country].filter(Boolean).join(', ');
  } catch (x) { console.error('ip-api: ' + x); }
  return '';
}
function abaIps_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_IPS);
  if (!sh) { sh = ss.insertSheet(ABA_IPS); sh.appendRow(CAB_IPS); sh.setFrozenRows(1); }
  return sh;
}
function linhasIps_(sh) {
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, CAB_IPS.length).getValues();
}
/* Anota o IP do e-mail e devolve quantos IPs diferentes ele já usou */
function registrarIp_(email, ip) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = abaIps_();
    const rows = linhasIps_(sh);
    let qtd = 0, achou = 0;
    rows.forEach((r, i) => {
      if (String(r[0]).toLowerCase() !== email) return;
      qtd++;
      if (String(r[1]) === ip) achou = i + 2;
    });
    if (achou) sh.getRange(achou, 4, 1, 2).setValues([[new Date(), (+rows[achou - 2][4] || 0) + 1]]);
    else { sh.appendRow([email, ip, new Date(), new Date(), 1, '']); qtd++; achou = sh.getLastRow(); }
    if (!achou || !rows[achou - 2] || !rows[achou - 2][5]) {
      const local = localIp_(ip);
      if (local) sh.getRange(achou, 6).setValue(local);
    }
    return qtd;
  } finally { lock.releaseLock(); }
}
function limparIps_(email) {
  const sh = abaIps_();
  const rows = linhasIps_(sh);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]).toLowerCase() === email) sh.deleteRow(i + 2);
  }
}
/* { email: [{ ip, local, primeiro, ultimo, acessos }, ...] } */
function ipsPorEmail_() {
  const m = {};
  const iso = x => x instanceof Date ? x.toISOString() : String(x || '');
  linhasIps_(abaIps_()).forEach(r => {
    const em = String(r[0]).toLowerCase();
    (m[em] = m[em] || []).push({ ip: String(r[1]), local: String(r[5] || ''), primeiro: iso(r[2]), ultimo: iso(r[3]), acessos: +r[4] || 0 });
  });
  return m;
}
/* Quantas contas entraram com 1 IP, com 2 e com mais de LIMITE_IPS */
function resumoIps_() {
  const m = ipsPorEmail_();
  const contas = Object.keys(m).map(em => ({ email: em, ips: m[em].length, lista: m[em] })).sort((a, b) => b.ips - a.ips);
  return {
    limite: LIMITE_IPS,
    contas: contas.length,
    umIp: contas.filter(c => c.ips === 1).length,
    variosIps: contas.filter(c => c.ips > 1).length,
    acimaDoLimite: contas.filter(c => c.ips > LIMITE_IPS).length,
    lista: contas
  };
}
/* Rode no editor para ver o resumo no registro de execução */
function verIps() { console.log(JSON.stringify(resumoIps_(), null, 2)); }

/* Área de membros -> planilha: ?email=fulano@x.com&ip=1.2.3.4 responde se tem acesso e se tem o bump */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const email = String(p.email || '').trim().toLowerCase();
  if (!email) return json_({ ok: true, servico: 'acesso avatar milionario' });
  const sh = aba_();
  const r = acharLinha_(sh, email);
  if (!r) return json_({ acesso: false, bump: false });
  const v = sh.getRange(r, 1, 1, COLS).getValues()[0];
  const on = x => String(x).trim().toLowerCase() === 'ativo';
  if (String(v[7]).indexOf(OBS_IPS) === 0) return json_({ acesso: false, bump: false, bloqueioIp: true });
  const ip = String(p.ip || '').trim();
  if (on(v[2]) && /^\d{1,3}(\.\d{1,3}){3}$/.test(ip)) {
    const qtd = registrarIp_(email, ip);
    /* Passou do limite: só marca o alerta (quem bloqueia é você, no painel) */
    if (qtd > LIMITE_IPS) {
      const obs = OBS_ALERTA + ' (' + qtd + ' IPs) em ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      if (!v[7] || String(v[7]).indexOf(OBS_ALERTA) === 0) salvar_(email, { observacao: obs });
    }
  }
  return json_({ acesso: on(v[2]), bump: on(v[3]), nome: v[1] || '' });
}
