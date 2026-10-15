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
const PRODUTO_BUMP = '';

const ABA = 'Alunos';
const CABECALHO = ['email', 'nome', 'acesso', 'bump', 'evento', 'pedido', 'atualizado em'];
const COLS = CABECALHO.length;

/* Status e eventos da Kiwify que liberam e que bloqueiam */
const LIBERA = ['paid', 'approved', 'order_approved', 'compra_aprovada'];
const BLOQUEIA = ['refunded', 'chargedback', 'chargeback', 'order_refunded', 'compra_reembolsada', 'subscription_canceled'];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(CABECALHO);
    sh.setFrozenRows(1);
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

/* Grava ou atualiza a linha do e-mail. Campos ausentes (undefined) mantêm o valor que já estava. */
function salvar_(email, campos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = aba_();
    const r = acharLinha_(sh, email);
    const atual = r ? sh.getRange(r, 1, 1, COLS).getValues()[0] : [email, '', '', '', '', '', ''];
    const pega = (v, i) => (v === undefined || v === '' ? atual[i] : v);
    const linha = [
      email,
      pega(campos.nome, 1),
      pega(campos.acesso, 2),
      pega(campos.bump, 3),
      pega(campos.evento, 4),
      pega(campos.pedido, 5),
      new Date()
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
  return sh.getRange(2, 1, n - 1, COLS).getValues().map(v => ({
    email: v[0], nome: v[1], acesso: v[2], bump: v[3], evento: v[4], pedido: v[5],
    quando: v[6] instanceof Date ? v[6].toISOString() : String(v[6])
  })).reverse();
}

/* Painel: { acao, senha, email?, nome? }
   acao: listar | liberar | bloquear | liberar_bump | bloquear_bump */
function painel_(d) {
  if (!SENHA_PAINEL || d.senha !== SENHA_PAINEL) {
    Utilities.sleep(1500);
    return json_({ ok: false, erro: 'senha' });
  }
  const email = String(d.email || '').trim().toLowerCase();
  const mudancas = {
    liberar: { acesso: 'ativo' },
    bloquear: { acesso: 'bloqueado' },
    liberar_bump: { bump: 'ativo' },
    bloquear_bump: { bump: 'bloqueado' }
  }[d.acao];
  if (mudancas) {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, erro: 'e-mail inválido' });
    mudancas.nome = String(d.nome || '').trim();
    mudancas.evento = 'manual';
    salvar_(email, mudancas);
  }
  if (d.acao === 'visitas') return json_({ ok: true, stats: estatisticas_() });
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

/* Evento enviado pelas páginas: { ev, pagina, afiliado, origem, visitante, visita, segundos, cliques } */
function registrar_(d) {
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
  if (!novo) return json_({ ok: true, ignorado: status || evento });

  /* Separa o que é bump do que é o curso */
  const prods = produtos_(d);
  const ehBump = p => PRODUTO_BUMP && p.toLowerCase().indexOf(PRODUTO_BUMP.toLowerCase()) >= 0;
  const temBump = prods.some(ehBump);
  const temCurso = !prods.length || prods.some(p => !ehBump(p));

  const nome = cli.full_name || cli.name || [cli.first_name, cli.last_name].filter(String).join(' ') || '';
  const campos = { nome: nome, evento: evento || status, pedido: d.order_id || d.order_ref || d.id || '' };
  if (temCurso) campos.acesso = novo;
  if (temBump) campos.bump = novo;
  salvar_(email, campos);
  return json_({ ok: true, email: email, acesso: campos.acesso || null, bump: campos.bump || null });
}

/* Área de membros -> planilha: ?email=fulano@x.com responde se tem acesso e se tem o bump */
function doGet(e) {
  const email = String((e && e.parameter && e.parameter.email) || '').trim().toLowerCase();
  if (!email) return json_({ ok: true, servico: 'acesso avatar milionario' });
  const sh = aba_();
  const r = acharLinha_(sh, email);
  if (!r) return json_({ acesso: false, bump: false });
  const v = sh.getRange(r, 1, 1, 4).getValues()[0];
  const on = x => String(x).trim().toLowerCase() === 'ativo';
  return json_({ acesso: on(v[2]), bump: on(v[3]), nome: v[1] || '' });
}
