/* Avatar Milionário - contador de visitas, cliques, tempo na página e cliques no checkout.
   A página define antes: window.AM_AFILIADO = 'nome-do-afiliado' (vazio na página principal). */
(function () {
  var API = 'https://script.google.com/macros/s/AKfycbyw-PCFL5Ed1hpDJp18ZlNqYBeUs3GD_D3gMnNuYYyof5XHwb54WJCqd70FAVkwBSL_/exec';
  if (!API) return;

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
