/* Avatar Milionário - contador de visitas, cliques, tempo na página e cliques no checkout.
   A página define antes: window.AM_AFILIADO = 'nome-do-afiliado' (vazio na página principal). */

/* Contagem regressiva na faixa do topo: a oferta "acaba" à meia-noite (horário de Brasília) e recomeça todo dia */
(function () {
  var faixa = document.getElementById('topIn');
  if (!faixa) return;
  var bloco = '<span>🔥 <b>Último dia</b></span><span>A oferta acaba em <b class="am-cd">--:--:--</b></span>';
  faixa.innerHTML = new Array(9).join(bloco);
  function tic() {
    var agora = new Date(), br = new Date(agora.getTime() + (agora.getTimezoneOffset() - 180) * 60000);
    var falta = 86400 - (br.getHours() * 3600 + br.getMinutes() * 60 + br.getSeconds());
    var t = [Math.floor(falta / 3600), Math.floor(falta % 3600 / 60), falta % 60].map(function (n) { return (n < 10 ? '0' : '') + n; }).join(':');
    [].forEach.call(document.querySelectorAll('.am-cd'), function (e) { e.textContent = t; });
  }
  tic();
  setInterval(tic, 1000);
})();

(function () {
  var API = 'https://script.google.com/macros/s/AKfycbyw-PCFL5Ed1hpDJp18ZlNqYBeUs3GD_D3gMnNuYYyof5XHwb54WJCqd70FAVkwBSL_/exec';
  /* Só conta visitas no site de verdade (ignora pré-visualizações e testes) */
  if (!API || !/(^|\.)avatarmilionario\.com$/.test(location.hostname)) return;

  function id() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }
  var visitante = '';
  try { visitante = localStorage.am_vis || (localStorage.am_vis = id()); } catch (e) { visitante = id(); }
  var visita = id();

  /* De onde veio: ?utm_source= ou ?src=, senão o site anterior, senão "direto" */
  var q = new URLSearchParams(location.search);
  var origem = q.get('utm_source') || q.get('src') || '';
  if (!origem && document.referrer) {
    try {
      var h = new URL(document.referrer).hostname.replace(/^www\./, '').replace(/^(l|lm|m)\./, '');
      if (h !== location.hostname.replace(/^www\./, '')) origem = h;
    } catch (e) {}
  }
  origem = origem || 'direto';

  var base = { pagina: location.pathname, afiliado: window.AM_AFILIADO || '', origem: origem, visitante: visitante, visita: visita };
  function envia(ev, extra) {
    var d = { ev: ev };
    for (var k in base) d[k] = base[k];
    for (var j in extra || {}) d[j] = extra[j];
    var corpo = JSON.stringify(d);
    if (navigator.sendBeacon && navigator.sendBeacon(API, new Blob([corpo], { type: 'text/plain' }))) return;
    try { fetch(API, { method: 'POST', body: corpo, keepalive: true, mode: 'no-cors' }); } catch (e) {}
  }

  envia('visita');

  /* Ao vivo: avisa a cada 45s que a página continua aberta (só enquanto está na tela) */
  function ping() { if (document.visibilityState === 'visible') envia('ping'); }
  ping();
  setInterval(ping, 45000);

  /* Leva a página e a origem para o checkout da Kiwify (aparece na venda: src e utm_source) */
  [].forEach.call(document.querySelectorAll('a[href*="pay.kiwify.com.br"]'), function (a) {
    try {
      var u = new URL(a.href);
      u.searchParams.set('src', (window.AM_AFILIADO || 'principal').toLowerCase());
      if (!u.searchParams.get('utm_source')) u.searchParams.set('utm_source', origem);
      ['utm_medium', 'utm_campaign', 'utm_content'].forEach(function (k) { if (q.get(k)) u.searchParams.set(k, q.get(k)); });
      a.href = u.toString();
    } catch (e) {}
  });

  /* Cliques na página e cliques que vão para o checkout */
  var cliques = 0, checkout = false;
  document.addEventListener('click', function (e) {
    cliques++;
    var a = e.target.closest && e.target.closest('a[href]');
    if (a && /pay\.kiwify\.com\.br|kiwify\.app/.test(a.href) && !checkout) { checkout = true; envia('checkout'); }
  }, true);

  /* Tempo com a página aberta e visível */
  var visivel = 0, desde = document.visibilityState === 'visible' ? Date.now() : 0, fechou = false;
  function pausa() { if (desde) { visivel += Date.now() - desde; desde = 0; } }
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'visible') desde = Date.now();
    else { pausa(); fim(); }
  });
  function fim() {
    if (fechou) return;
    fechou = true;
    envia('fim', { segundos: Math.round(visivel / 1000), cliques: cliques });
  }
  addEventListener('pagehide', function () { pausa(); fim(); });
})();
