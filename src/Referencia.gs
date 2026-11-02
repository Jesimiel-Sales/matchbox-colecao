/**
 * Base de referência e fotos por variação.
 *
 * Fontes (consultadas sob demanda e guardadas em cache na aba _Ref):
 *  - fcarnahan.com: lista de variações das rodas regulares segundo Charlie Mack
 *    (publicada com permissão do autor) e páginas com foto por variação.
 *  - Arquivo Nick Jones / Lesney Matchbox (vbd2-archive.in-nz.com): páginas por modelo,
 *    Lesney e Superfast, com fotos legendadas de cada variação.
 * As fotos são exibidas a partir do site de origem, com crédito e link.
 */

const REF = {
  FC_LIST: 'https://www.fcarnahan.com/pg/regvaria.html',
  fcVar: (n, l, v) => 'https://www.fcarnahan.com/rw' + (Math.floor(n / 10) * 10 + 9) + '/' + n + l.toLowerCase() + v + '.html',
  njRW: (n, l) => 'https://vbd2-archive.in-nz.com/1index/xx' + n + l.toLowerCase() + '.htm',
  njSF: (n, l) => 'https://vbd2-archive.in-nz.com/1index/sf' + n + '-' + l.toLowerCase() + '.htm',
  LETRAS_SF: 'abcdefghijklmnop'.split(''),
  UA: { 'User-Agent': 'Mozilla/5.0 (colecao-matchbox; uso pessoal)' }
};

/* Durante o uso do app a referência é só LOCAL (aba _Ref + banco de imagens no Drive).
   Os sites só são consultados pela importação em segundo plano (montarReferenciaLocal). */
let MODO_ONLINE = false;

/* ---------------------------------------------------------------- cache em planilha */

function refSheet_() { return refAba_('_Ref'); }

function refGet_(key) {
  const c = CacheService.getScriptCache().get('ref:' + key);
  if (c) return JSON.parse(c);
  const sh = refSheet_(), last = sh.getLastRow();
  if (last < 2) return null;
  const cell = sh.getRange(2, 1, last - 1, 1).createTextFinder(key).matchEntireCell(true).findNext();
  if (!cell) return null;
  const txt = sh.getRange(cell.getRow(), 3).getValue();
  try { CacheService.getScriptCache().put('ref:' + key, txt, 21600); } catch (e) {}
  return JSON.parse(txt);
}

function refPut_(key, obj) {
  let txt = JSON.stringify(obj);
  if (txt.length > 49000) txt = JSON.stringify(Object.assign({}, obj, { texto: '' })).slice(0, 49000);
  const sh = refSheet_(), last = sh.getLastRow();
  const cell = last >= 2 ? sh.getRange(2, 1, last - 1, 1).createTextFinder(key).matchEntireCell(true).findNext() : null;
  if (cell) sh.getRange(cell.getRow(), 2, 1, 2).setValues([[new Date(), txt]]);
  else sh.appendRow([key, new Date(), txt]);
  try { CacheService.getScriptCache().put('ref:' + key, txt, 21600); } catch (e) {}
}

/* ---------------------------------------------------------------- HTML utilitários */

function decode_(s) {
  const nomes = { nbsp: ' ', amp: '&', quot: '"', apos: "'", lt: '<', gt: '>', copy: '©', bull: '·', middot: '·',
    ndash: '–', mdash: '—', raquo: '»', laquo: '«', hellip: '…', deg: '°', frac12: '½', reg: '®', trade: '™' };
  return String(s || '')
    .replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (m, d) => String.fromCharCode(+d))
    .replace(/&([a-z0-9]+);/gi, (m, n) => (n.toLowerCase() in nomes ? nomes[n.toLowerCase()] : ' '));
}
function strip_(html) {
  return decode_(String(html || '').replace(/<script[\s\S]*?(<\/script>|$)|<style[\s\S]*?(<\/style>|$)/gi, ' ')
    .replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
}
function resolve_(base, src) {
  src = decode_(src).trim();
  if (/^https?:\/\//i.test(src)) return src;
  const origin = base.match(/^https?:\/\/[^/]+/)[0];
  if (src.indexOf('//') === 0) return 'https:' + src;
  if (src.charAt(0) === '/') return origin + src;
  const parts = base.replace(/[?#].*$/, '').split('/'); parts.pop();
  src.split('/').forEach(p => { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); });
  return parts.join('/');
}

const IGNORAR_IMG = /(\/pre\/|\/images\/(back|next|home)|button|header|banner|roads\.|vvvvvv|icon|logo|counter|spacer|bookad|forum|home1|\/mbx\.jpg|\.gif($|\?))/i;
const VERSAO_CACHE = 'p2:'; // muda quando a leitura das páginas muda (invalida o cache antigo)

/** Trecho de HTML -> texto, descartando pedaços de tag cortados nas pontas. */
function trecho_(html) {
  return strip_(String(html || '').replace(/^[^<]*?>/, ' ').replace(/<[^>]*$/, ' '));
}
/** Descarta "legendas" que na verdade são restos de código da página. */
function legendaValida_(s) {
  s = String(s || '').replace(/©.*$/, '').replace(/\s+/g, ' ').trim();
  if (s.length < 4) return '';
  if (/[{};]|=\s*["']|\\x3C|bgcolor|cellpadding|google_|document\.|function\s*\(|var\s+\w+\s*=/i.test(s)) return '';
  return s;
}

/** Extrai título, texto e imagens (com a legenda que vem logo após cada imagem). */
function parsePage_(html, url) {
  html = String(html || '').replace(/<!--[\s\S]*?-->/g, ' ').replace(/<script[\s\S]*?(<\/script>|$)/gi, ' ')
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, ' ').replace(/<style[\s\S]*?<\/style>/gi, ' ');
  const title = strip_((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || '');
  const body = (html.match(/<body[\s\S]*<\/body>/i) || [html])[0];
  const tags = [];
  const re = /<img\b[^>]*>/gi; let m;
  while ((m = re.exec(body))) tags.push({ tag: m[0], start: m.index, end: m.index + m[0].length });
  const imgs = [], vistos = {};
  tags.forEach((t, i) => {
    const src0 = (t.tag.match(/\bsrc\s*=\s*["']?([^"'\s>]+)/i) || [])[1];
    if (!src0) return;
    const src = resolve_(url, src0);
    if (IGNORAR_IMG.test(src) || vistos[src]) return;
    vistos[src] = true;
    const alt = decode_((t.tag.match(/\balt\s*=\s*["']([^"']*)["']/i) || [])[1] || '').trim();
    const next = i + 1 < tags.length ? tags[i + 1].start : Math.min(body.length, t.end + 1500);
    let cap = legendaValida_(trecho_(body.slice(t.end, next)).slice(0, 320));
    if (!cap) cap = legendaValida_(trecho_(body.slice(Math.max(0, t.start - 400), t.start)).slice(-200));
    if (alt && !/^photo\s|\.jpe?g$/i.test(alt) && cap.indexOf(alt) < 0) cap = (alt + (cap ? ' — ' + cap : '')).trim();
    imgs.push({ src: src, cap: cap });
  });
  return { ok: true, title: title, imgs: imgs.slice(0, 40), texto: strip_(body).slice(0, 4000) };
}

function fetchParsed_(url) {
  const hit = refGet_(VERSAO_CACHE + url);
  if (hit) return hit;
  if (!MODO_ONLINE) return { ok: false, local: true };
  let res;
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: REF.UA });
    res = r.getResponseCode() === 200 ? parsePage_(r.getContentText(), url) : { ok: false, code: r.getResponseCode() };
  } catch (e) { res = { ok: false, erro: String(e) }; }
  if (res.ok || res.code === 404) refPut_(VERSAO_CACHE + url, res);
  return res;
}

function fetchAllParsed_(urls) {
  const out = {}, falta = [];
  urls.forEach(u => { const h = refGet_(VERSAO_CACHE + u); if (h) out[u] = h; else falta.push(u); });
  if (!MODO_ONLINE) { falta.forEach(u => out[u] = { ok: false, local: true }); return out; }
  if (falta.length) {
    let resps = [];
    try { resps = UrlFetchApp.fetchAll(falta.map(u => ({ url: u, muteHttpExceptions: true, followRedirects: true, headers: REF.UA }))); }
    catch (e) { falta.forEach(u => out[u] = { ok: false, erro: String(e) }); return out; }
    resps.forEach((r, i) => {
      const u = falta[i];
      const res = r.getResponseCode() === 200 ? parsePage_(r.getContentText(), u) : { ok: false, code: r.getResponseCode() };
      if (res.ok || res.code === 404) refPut_(VERSAO_CACHE + u, res);
      out[u] = res;
    });
  }
  return out;
}

/* ---------------------------------------------------------------- lista Charlie Mack (rodas regulares) */

function fcModelo_(n, l) {
  const key = 'fc:' + n + l.toUpperCase();
  let m = refGet_(key);
  if (m) return m;
  if (refGet_('fc:carregado') || !MODO_ONLINE) return null;
  carregarListaMack_();
  return refGet_(key);
}

function carregarListaMack_() {
  const r = UrlFetchApp.fetch(REF.FC_LIST, { muteHttpExceptions: true, headers: REF.UA });
  if (r.getResponseCode() !== 200) throw new Error('Lista Charlie Mack indisponível (' + r.getResponseCode() + ')');
  const html = r.getContentText();
  const linhas = decode_(html.replace(/<li\b[^>]*>/gi, '\n').replace(/<\/?(ul|ol|p|div|tr|h\d)\b[^>]*>/gi, '\n')
    .replace(/<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>/gi, ' [[$1]] ').replace(/<[^>]+>/g, ' '))
    .split('\n').map(s => s.replace(/\s+/g, ' ').trim()).filter(String);
  const modelos = {}; let atual = null;
  linhas.forEach(s => {
    const mm = s.match(/^(\d{1,2})\s*([A-H])\s+(.+?),?\s+Issued\s+(\d{4})/i);
    if (mm) {
      atual = { cod: mm[1] + mm[2].toUpperCase(), nome: mm[3].replace(/\[\[[^\]]+\]\]/g, '').trim(), ano: +mm[4], vars: {} };
      modelos[atual.cod] = atual; return;
    }
    const vv = atual && s.match(/^(?:\[\[([^\]]+)\]\]\s*)?(\d{1,2})\.\s+(.+)$/);
    if (vv) {
      const link = vv[1] || (s.match(/\[\[([^\]]+)\]\]/) || [])[1] || '';
      atual.vars[vv[2]] = { desc: vv[3].replace(/\[\[[^\]]+\]\]/g, '').trim(), link: link ? resolve_(REF.FC_LIST, link) : '' };
    }
  });
  // grava tudo de uma vez (antes eram 232 gravações, uma por modelo, e a 1ª abertura levava ~2 min)
  const sh = refSheet_(), agora = new Date(), last = sh.getLastRow();
  const existentes = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach((r, i) => existentes[r[0]] = i + 2);
  const novas = [];
  Object.keys(modelos).forEach(k => {
    const linha = ['fc:' + k, agora, JSON.stringify(modelos[k])];
    if (existentes[linha[0]]) sh.getRange(existentes[linha[0]], 1, 1, 3).setValues([linha]); else novas.push(linha);
  });
  if (novas.length) sh.getRange(sh.getLastRow() + 1, 1, novas.length, 3).setValues(novas);
  refPut_('fc:carregado', { total: Object.keys(modelos).length, em: agora.toISOString() });
}

/* ---------------------------------------------------------------- correspondência PT → EN */

const PT_EN = {
  verde: 'green', vermelho: 'red', vermelha: 'red', amarelo: 'yellow', amarela: 'yellow', azul: 'blue', laranja: 'orange',
  branco: 'white', branca: 'white', preto: 'black', preta: 'black', dourado: 'gold', dourada: 'gold', prata: 'silver',
  prateado: 'silver', prateada: 'silver', cinza: 'grey|gray', marrom: 'brown', bege: 'beige|tan', creme: 'cream',
  rosa: 'pink', roxo: 'purple', roxa: 'purple', 'lilás': 'lilac|purple', vinho: 'maroon', oliva: 'olive', 'bordô': 'maroon',
  escuro: 'dark', escura: 'dark', 'esc.': 'dark', claro: 'light', clara: 'light', 'cl.': 'light', 'met.': 'metal',
  'metálica': 'metallic|metal', 'metálico': 'metallic', 'pl.': 'plastic', cromada: 'chrome', cromado: 'chrome',
  capota: 'canopy', 'caçamba': 'tipper|dump', esteira: 'tracks|treads', rodas: 'wheels', roda: 'wheel|roller',
  'transpar.': 'clear', 'âmbar': 'amber', turquesa: 'turquoise', marfim: 'ivory|cream', caramelo: 'tan', 'limão': 'lemon',
  raiada: 'spoke', motorista: 'driver', guindaste: 'crane', 'decalque': 'decal', 'd.': 'decal', 'ad.': 'label',
  'escada': 'ladder', 'trailer': 'trailer', 'azul-claro': 'light blue', 'verde-claro': 'light green'
};

/** Palavras-chave (em inglês) que descrevem a variação, com peso. */
function termosItem_(item) {
  const termos = [];
  const add = (txt, peso) => {
    String(txt || '').toLowerCase().split(/[\s,/;()]+/).forEach((w, i, arr) => {
      if (!w) return;
      if (/^\d+$/.test(w) && arr[i + 1] === 'mm') { termos.push({ re: w + '\\s?mm', peso: peso + 1, w: w + 'mm' }); return; }
      const en = PT_EN[w];
      if (en) en.split('|').forEach(e => termos.push({ re: '\\b' + e, peso: peso, w: e }));
      else if (w.length >= 4 && /^[a-z0-9]+$/.test(w)) termos.push({ re: '\\b' + w, peso: peso, w: w }); // nomes próprios / marcas (ex.: BP, Esso)
    });
  };
  add(item.cor, 3); add(item.roda, 2); add(item.base, 1); add(item.interior, 1); add(item.vidro, 1);
  add(item.outros, 2); add(item.adesivo, 2);
  return termos;
}

function pontuar_(caption, termos) {
  const c = String(caption || '').toLowerCase(); let score = 0; const marcadas = [];
  termos.forEach(t => { if (new RegExp(t.re, 'i').test(c)) { score += t.peso; if (marcadas.indexOf(t.w) < 0) marcadas.push(t.w); } });
  return { score: score, marcadas: marcadas };
}

function semelhancaNome_(a, b) {
  const tok = s => String(s || '').toLowerCase().replace(/matchbox|superfast|lesney|\bmb\d+[-\w]*/g, ' ')
    .split(/[^a-z0-9]+/).filter(w => w.length > 1);
  const A = tok(a), B = tok(b);
  if (!A.length || !B.length) return 0;
  return A.filter(w => B.indexOf(w) >= 0).length / Math.max(A.length, 1);
}

function parseCodigo_(s) {
  const m = String(s || '').trim().match(/^0*(\d{1,3})\s*([A-Za-z])\s*(\d+)?/);
  return m ? { n: +m[1], l: m[2].toUpperCase(), v: m[3] ? +m[3] : null } : null;
}

/* ---------------------------------------------------------------- API usada pelo app */

/**
 * Referência para um item (existente ou em cadastro): nome/ano do modelo, lista de variações
 * do Charlie Mack (Lesney) e fotos ordenadas pela aderência à variação.
 */
function getReference(name, item) {
  const out = { modelo: null, variacoes: [], fotos: [], links: [], avisos: [] };
  // Lesney: o código Charlie Mack é a classificação detalhada (mack2); a coluna Cl. (mack) só entra se ela faltar
  const cod = name === 'Lesney' ? (parseCodigo_(item.mack2) || parseCodigo_(item.mack)) : (parseCodigo_(item.mack) || parseCodigo_(item.mack2));
  if (!cod) { out.avisos.push('Código Mack não reconhecido.'); return out; }
  const termos = termosItem_(item);
  const addFotos = (pag, url, fonte) => {
    if (!pag || !pag.ok) return;
    out.links.push({ titulo: fonte + (pag.title ? ' — ' + pag.title : ''), url: url });
    pag.imgs.forEach(im => {
      const p = pontuar_(im.cap, termos);
      out.fotos.push({ url: im.src, drive: im.drive || null, legenda: im.cap, fonte: fonte, pagina: url, score: p.score, marcadas: p.marcadas });
    });
  };

  if (name === 'Lesney') {
    let fcUrl = null;
    try {
      const mod = fcModelo_(cod.n, cod.l);
      if (mod) {
        out.modelo = { nome: mod.nome, ano: mod.ano, fonte: 'Charlie Mack (via fcarnahan.com)' };
        out.variacoes = Object.keys(mod.vars).map(v => ({ v: +v, desc: mod.vars[v].desc, link: mod.vars[v].link, atual: +v === cod.v }));
        if (cod.v && mod.vars[cod.v] && mod.vars[cod.v].link) fcUrl = mod.vars[cod.v].link;
      }
    } catch (e) { out.avisos.push('Lista Charlie Mack: ' + e.message); }
    const nj = REF.njRW(cod.n, cod.l);
    const pags = fetchAllParsed_(fcUrl ? [fcUrl, nj] : [nj]);
    if (fcUrl) addFotos(pags[fcUrl], fcUrl, 'fcarnahan.com (variação ' + cod.n + cod.l + cod.v + ')');
    addFotos(pags[nj], nj, 'Nick Jones / Lesney Matchbox');
    // foto da página da variação do Mack é a mais específica: sobe na ordenação
    out.fotos.forEach(f => { if (f.fonte.indexOf('fcarnahan') === 0) f.score += 5; });
  } else {
    const url = superfastPagina_(cod.n, item.nome);
    if (url) {
      const pag = fetchParsed_(url);
      addFotos(pag, url, 'Nick Jones / Lesney Matchbox');
      if (pag && pag.ok) {
        const nome = pag.title.replace(/.*?MB\s*\d+\s*-?\s*[a-z]\s*/i, '').trim();
        out.modelo = { nome: nome || pag.title, ano: null, fonte: 'Nick Jones / Lesney Matchbox' };
      }
    } else out.avisos.push('Sem fotos de referência deste modelo no arquivo importado (as fotos do BAMCA entram quando a cópia for autorizada).');
  }
  out.fotos.sort((a, b) => b.score - a.score);
  if (out.fotos.length) out.fotos[0].melhor = true;
  if (!MODO_ONLINE && !out.fotos.length && !out.variacoes.length && !out.modelo)
    out.avisos.push('A referência deste modelo ainda não foi importada. Ela entra na próxima importação (menu da planilha > Matchbox App > Importar referência e fotos).');
  return out;
}

/** Versão chamada pelo app (texto JSON). */
function getReferenceJson(name, item) { return JSON.stringify(getReference(name, item)); }

/** Descobre a página Superfast do arquivo Nick Jones para um número + nome de modelo. */
function superfastPagina_(n, nome) {
  const chaveIdx = 'njsf:' + n;
  let idx = refGet_(chaveIdx);
  if (!idx && !MODO_ONLINE) return null;
  if (!idx) {
    const urls = REF.LETRAS_SF.map(l => REF.njSF(n, l));
    const pags = fetchAllParsed_(urls);
    idx = {};
    REF.LETRAS_SF.forEach((l, i) => { const p = pags[urls[i]]; if (p && p.ok) idx[l] = p.title; });
    refPut_(chaveIdx, idx);
  }
  // Só usa a página se o nome do modelo bater (pelo menos metade das palavras). Antes, quando o número
  // tinha uma única página, ela era usada para todos os modelos daquele número (ex.: fotos do 1A Mercedes Truck no 1C Dodge Challenger).
  let melhor = null, nota = 0;
  Object.keys(idx).forEach(l => { const s = semelhancaNome_(nome, idx[l]); if (s > nota) { nota = s; melhor = l; } });
  return melhor && nota >= 0.5 ? REF.njSF(n, melhor) : null;
}

/** Usado quando o navegador não consegue exibir a imagem direto do site de origem. */
function getImageData(url) {
  const cache = CacheService.getScriptCache(), key = 'img:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, url));
  const hit = cache.get(key);
  if (hit) return hit;
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: REF.UA });
  if (r.getResponseCode() !== 200) throw new Error('Imagem indisponível (' + r.getResponseCode() + ')');
  const b = r.getBlob();
  const data = 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
  if (data.length < 95000) cache.put(key, data, 21600);
  return data;
}

/* ---------------------------------------------------------------- banco de imagens (Drive) */

/** Copia uma foto de referência para a pasta "Matchbox • Banco de imagens" (uma vez por URL)
 *  e devolve o ID do arquivo. Os arquivos ficam visíveis só para quem tem o link, o que permite
 *  ao Google entregar miniaturas redimensionadas e rápidas. */
function bancoImagem_(url) {
  const key = 'bk:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, url));
  const hit = refGet_(key);
  if (hit && hit.id) return hit.id;
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: REF.UA });
  if (r.getResponseCode() !== 200) throw new Error('imagem indisponível (' + r.getResponseCode() + ')');
  const b = r.getBlob(), ct = b.getContentType() || '';
  if (!/^image\//.test(ct)) throw new Error('o endereço não é uma imagem');
  b.setName(key.slice(3) + '.' + (ct.split('/')[1] || 'jpg').replace('jpeg', 'jpg'));
  const f = folder_('pasta_banco_id', 'Matchbox • Banco de imagens').createFile(b);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* conta sem permissão de compartilhar: segue privado */ }
  f.setDescription('Fonte: ' + url);
  refPut_(key, { id: f.getId(), url: url });
  return f.getId();
}

/* ---------------------------------------------------------------- importação única da referência */

function iniciarReferenciaLocal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'montarReferenciaLocal').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('montarReferenciaLocal').timeBased().after(10 * 1000).create();
  avisar_('Importação iniciada em segundo plano. Ela roda sozinha, em etapas, e pode levar algumas horas. O andamento aparece no app em Menu > Segurança dos dados.');
  return 'iniciado';
}

/** Lista de páginas de referência a importar (todas as rodas regulares do Mack + Superfast 1-75). */
function tarefasReferencia_() {
  const tarefas = [];
  const sh = refSheet_(), last = sh.getLastRow();
  const chaves = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  chaves.filter(k => /^fc:\d+[A-Z]$/.test(k)).forEach(k => {
    const cod = k.slice(3), n = +cod.slice(0, -1), l = cod.slice(-1);
    tarefas.push(REF.njRW(n, l));
    const mod = refGet_(k);
    if (mod) Object.keys(mod.vars).forEach(v => { if (mod.vars[v].link) tarefas.push(mod.vars[v].link); });
  });
  for (let n = 1; n <= 75; n++) {
    const idx = refGet_('njsf:' + n);
    if (idx) Object.keys(idx).forEach(l => tarefas.push(REF.njSF(n, l)));
  }
  return tarefas;
}

/** Importa a referência uma única vez: lista Mack, páginas por modelo e TODAS as fotos (para o banco
 *  de imagens no Drive). Depois disso o app não precisa acessar nenhum site. Roda em etapas de ~5 min. */
function montarReferenciaLocal() {
  MODO_ONLINE = true;
  const inicio = Date.now(), limite = 4.5 * 60 * 1000, props = PropertiesService.getScriptProperties();
  const tempo = () => Date.now() - inicio > limite;
  const st = { fase: '', feitos: 0, total: 0, fotos: 0, em: new Date().toISOString() };
  const salvar = () => props.setProperty('ref_local', JSON.stringify(st));
  const reagendar = () => {
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'montarReferenciaLocal').forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('montarReferenciaLocal').timeBased().after(60 * 1000).create();
  };
  // fase 1: lista Mack e índice das páginas Superfast
  st.fase = '1/3 lista Charlie Mack e índice Superfast';
  if (!refGet_('fc:carregado')) carregarListaMack_();
  for (let n = 1; n <= 75; n++) {
    if (!refGet_('njsf:' + n)) { superfastPagina_(n, ''); if (tempo()) { salvar(); reagendar(); return st; } }
  }
  // fase 2: páginas e fotos
  st.fase = '2/3 páginas e fotos de referência';
  const tarefas = tarefasReferencia_();
  st.total = tarefas.length;
  for (let i = 0; i < tarefas.length; i++) {
    const url = tarefas[i];
    const pag = fetchParsed_(url);
    if (pag && pag.ok && !pag.banco) {
      pag.imgs.forEach(im => {
        if (im.drive || tempo()) return;
        try { im.drive = bancoImagem_(im.src); st.fotos++; } catch (e) { im.drive = null; }
      });
      if (pag.imgs.every(im => im.drive !== undefined)) pag.banco = true;
      refPut_(VERSAO_CACHE + url, pag);
    }
    st.feitos = i + 1;
    if (tempo()) { salvar(); reagendar(); return st; }
  }
  // fase 3: foto principal de cada item (só local)
  st.fase = '3/3 foto principal de cada item';
  MODO_ONLINE = false;
  for (const name of Object.keys(SCHEMA)) {
    const sh = sheet_(name), last = lastDataRow_(sh);
    if (last < APP.FIRST) continue;
    const vals = sh.getRange(APP.FIRST, 1, last - APP.FIRST + 1, width_()).getValues();
    for (let i = 0; i < vals.length; i++) {
      const item = rowToItem_(name, vals[i]);
      if (!item.id || /^(var:)?drive:/.test(item.fotoRef)) continue;
      if (tempo()) { salvar(); MODO_ONLINE = true; reagendar(); return st; }
      const ref = getReference(name, item);
      const f = ref.fotos.find(x => x.drive);
      if (f) sh.getRange(APP.FIRST + i, APP.TECH_COL + 1).setValue('drive:' + f.drive + '|' + f.url);
    }
  }
  st.fase = 'concluída'; st.em = new Date().toISOString();
  salvar();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'montarReferenciaLocal').forEach(t => ScriptApp.deleteTrigger(t));
  return st;
}

/* ---------------------------------------------------------------- lista de procura */

/**
 * Todos os modelos Lesney da lista do Charlie Mack (importada na _Ref), em formato compacto,
 * mais a lista de procura que o Marco marcou. Lê só o trecho da _Ref com as chaves fc:.
 */
function getListaMack() {
  const sh = refSheet_(), last = sh.getLastRow(), modelos = {};
  if (last >= 2) {
    const chaves = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
    let ini = -1, fim = -1;
    chaves.forEach((k, i) => { if (/^fc:\d+[A-Z]$/.test(k)) { if (ini < 0) ini = i; fim = i; } });
    if (ini >= 0) {
      const textos = sh.getRange(2 + ini, 3, fim - ini + 1, 1).getValues();
      for (let i = ini; i <= fim; i++) {
        if (!/^fc:\d+[A-Z]$/.test(chaves[i])) continue;
        try {
          const m = JSON.parse(textos[i - ini][0]);
          modelos[m.cod] = { n: m.nome, a: m.ano, v: Object.keys(m.vars).map(k => [+k, m.vars[k].desc]).sort((x, y) => x[0] - y[0]) };
        } catch (e) { /* linha corrompida: ignora */ }
      }
    }
  }
  return JSON.stringify({ modelos: modelos, procura: lerProcura_() });
}

function lerProcura_() {
  try { return JSON.parse(cfgGet_('lista_procura') || '[]'); } catch (e) { return []; }
}

/** Grava a lista de procura inteira (vem do app). Cada item: { c: código Mack ou '', t: texto, em: data }. */
function setProcura(lista) {
  const limpa = (Array.isArray(lista) ? lista : []).slice(0, 2000).map(x => ({
    c: String(x.c || '').slice(0, 12), t: String(x.t || '').slice(0, 300), em: String(x.em || '').slice(0, 10), a: String(x.a || '').slice(0, 12)
  }));
  return withLock_(() => { cfgSet_('lista_procura', JSON.stringify(limpa)); return JSON.stringify(limpa); });
}
/* ================================================================ BAMCA (referência comparativa)
   Importação única, autorizada por Dean Dierschow (BAMCA) em 01/11/2026: lê a cópia do banco de dados
   publicada no GitHub (não acessa o site do BAMCA) e grava na aba _BAMCA só os modelos Lesney (Regular Wheels)
   e Superfast até 1983, indexados pelo código de modelo Charlie Mack. */
const BAMCA = {
  URL: 'https://raw.githubusercontent.com/ddierschow/bamca/master/sql/bamca.mysql.20260725',
  ABA: '_BAMCA',
  CAB: ['Chave', 'Casting', 'Modelo', 'Ano modelo', 'Var.', 'Carroceria', 'Base', 'Vidros', 'Interior', 'Rodas', 'Decoração', 'Fabricação', 'Data', 'Nota']
};

/** Lê as linhas de uma tabela do arquivo SQL (formato mysqldump). Devolve objetos {coluna: valor}. */
function tabelaSql_(txt, tabela) {
  const ini = txt.indexOf('CREATE TABLE `' + tabela + '` (');
  if (ini < 0) throw new Error('Tabela ' + tabela + ' não encontrada no arquivo.');
  const fimDef = txt.indexOf('\n) ENGINE', ini);
  const cols = (txt.slice(ini, fimDef).match(/^\s+`\w+`/gm) || []).map(c => c.trim().replace(/`/g, ''));
  const out = [], marca = 'INSERT INTO `' + tabela + '` VALUES';
  let p = txt.indexOf(marca, fimDef);
  while (p >= 0) {
    let i = p + marca.length;
    const n = txt.length;
    for (;;) {
      while (i < n && txt[i] !== '(' && txt[i] !== ';') i++;
      if (i >= n || txt[i] === ';') break;
      i++;
      const row = [];
      for (;;) {
        if (txt[i] === "'") {
          let j = i + 1, buf = '';
          for (;;) {
            const c = txt[j];
            if (c === '\\') { const e = txt[j + 1]; buf += e === 'n' ? '\n' : e === 'r' ? '' : e === 't' ? ' ' : e === '0' ? '' : e; j += 2; continue; }
            if (c === "'") { if (txt[j + 1] === "'") { buf += "'"; j += 2; continue; } break; }
            buf += c; j++;
          }
          row.push(buf); i = j + 1;
        } else {
          let j = i; while (txt[j] !== ',' && txt[j] !== ')') j++;
          const v = txt.slice(i, j); row.push(v === 'NULL' ? '' : v); i = j;
        }
        if (txt[i] === ',') { i++; continue; }
        i++; break; // ')'
      }
      const o = {}; cols.forEach((c, k) => { o[c] = row[k] == null ? '' : row[k]; }); out.push(o);
      while (i < n && (txt[i] === ',' || txt[i] === '\n' || txt[i] === ' ')) i++;
    }
    p = txt.indexOf(marca, i);
  }
  return out;
}

/** Monta as linhas da aba _BAMCA a partir do texto do arquivo SQL. (Separada para poder ser testada.) */
function linhasBamca_(txt) {
  const cast = {}; tabelaSql_(txt, 'casting').forEach(c => { cast[c.id] = c; });
  const porMod = {}; tabelaSql_(txt, 'variation').forEach(v => { (porMod[v.mod_id] = porMod[v.mod_id] || []).push(v); });
  const alias = {}; tabelaSql_(txt, 'alias').forEach(a => { if (a.type === 'mack') alias[a.id.toLowerCase()] = a.ref_id; });
  const linhas = [];
  [['L', 'RW'], ['S', 'SF']].forEach(([s, pre]) => {
    for (let n = 1; n <= 75; n++) for (const l of 'abcdefghijklmnopqrst') {
      const base = pre + ('0' + n).slice(-2) + l;
      const cid = [base, alias[base.toLowerCase()]].find(x => x && cast[x] && porMod[x]);
      if (!cid) continue;
      const c = cast[cid], ano = parseInt(c.first_year, 10) || 0;
      if (s === 'S' && (ano < 1969 || ano > 1983)) continue;
      const chave = s + n + l.toUpperCase();
      porMod[cid].filter(v => /^\w/.test(v.var)).sort((a, b) => a.var < b.var ? -1 : a.var > b.var ? 1 : 0).forEach(v => {
        linhas.push([chave, cid, String(c.rawname || '').replace(/;/g, ' '), c.first_year, v.var, v.body, v.base, v.windows, v.interior,
          v.wheels, v.deco, v.manufacture, v.date, v.note].map(x => String(x == null ? '' : x)));
      });
    }
  });
  return linhas;
}

/** Menu da planilha / editor: importa (ou reimporta) a referência do BAMCA. Leva cerca de 1 a 2 minutos. */
function importarBamca() {
  const r = UrlFetchApp.fetch(BAMCA.URL, { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Não consegui baixar o arquivo do BAMCA (HTTP ' + r.getResponseCode() + ').');
  const linhas = linhasBamca_(r.getContentText('ISO-8859-1'));
  if (linhas.length < 1000) throw new Error('Arquivo do BAMCA veio incompleto (' + linhas.length + ' variações).');
  const ss = refSs_();
  let sh = ss.getSheetByName(BAMCA.ABA);
  if (!sh) { sh = ss.insertSheet(BAMCA.ABA); sh.hideSheet(); }
  sh.clearContents();
  const W = BAMCA.CAB.length;
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  if (sh.getMaxRows() < linhas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), linhas.length + 1 - sh.getMaxRows());
  sh.getRange(1, 1, 1, W).setValues([BAMCA.CAB]).setFontWeight('bold');
  sh.getRange(2, 1, linhas.length, W).setNumberFormat('@');
  for (let i = 0; i < linhas.length; i += 4000) sh.getRange(2 + i, 1, Math.min(4000, linhas.length - i), W).setValues(linhas.slice(i, i + 4000));
  const modelos = new Set(linhas.map(l => l[0])).size;
  cfgSet_('bamca_importado', Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm') + ' · ' + modelos + ' modelos · ' + linhas.length + ' variações · fonte ' + BAMCA.URL.split('/').pop());
  CacheService.getScriptCache().remove('bamca_idx');
  Logger.log('BAMCA importado: ' + modelos + ' modelos, ' + linhas.length + ' variações.');
  return modelos + ' modelos, ' + linhas.length + ' variações';
}

/** Índice chave -> [primeira linha, quantidade], guardado em cache por 6 h. */
function indiceBamca_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('bamca_idx');
  if (hit) return JSON.parse(hit);
  const sh = refAba_(BAMCA.ABA);
  if (!sh || sh.getLastRow() < 2) return {};
  const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues(), idx = {};
  col.forEach((r, i) => { const k = r[0]; if (!k) return; if (!idx[k]) idx[k] = [i + 2, 0]; idx[k][1]++; });
  cache.put('bamca_idx', JSON.stringify(idx), 21600);
  return idx;
}

/** Variações do BAMCA para o modelo Charlie Mack do item (ex.: Lesney "01A2" -> L1A). JSON para o app. */
function getBamca(name, codigo) {
  const p = parseCodigo_(codigo);
  if (!p) return JSON.stringify({ ok: false });
  const chave = (name === 'Lesney' ? 'L' : 'S') + p.n + p.l, ix = indiceBamca_()[chave];
  if (!ix) return JSON.stringify({ ok: false, chave: chave });
  const vals = refAba_(BAMCA.ABA).getRange(ix[0], 1, ix[1], BAMCA.CAB.length).getDisplayValues();
  return JSON.stringify({
    ok: true, chave: chave, casting: vals[0][1], modelo: vals[0][2], ano: vals[0][3],
    v: vals.map(r => ({ var: r[4], body: r[5], base: r[6], windows: r[7], interior: r[8], wheels: r[9], deco: r[10], fab: r[11], data: r[12], obs: r[13] }))
  });
}

/** Teste pelo editor: mostra o que o app recebe para um código. */
function testarBamca() { Logger.log(getBamca('Lesney', '01A3').slice(0, 400)); Logger.log(getBamca('Superfast', '1A2').slice(0, 300)); }


/** Pelo editor: quantos itens Superfast recebiam fotos de outro modelo com a regra antiga. */
function auditarFotosSuperfast() {
  const d = JSON.parse(getData()), keys = d.schema.Superfast.keys, iN = keys.indexOf('nome'), iM = keys.indexOf('mack');
  let antes = 0, agora = 0, errados = 0; const ex = [];
  d.sheets.Superfast.forEach(r => {
    const c = parseCodigo_(r[iM]); if (!c) return;
    const idx = refGet_('njsf:' + c.n) || {}, ls = Object.keys(idx);
    let melhor = null, nota = 0;
    ls.forEach(l => { const s = semelhancaNome_(r[iN], idx[l]); if (s > nota) { nota = s; melhor = l; } });
    const velho = melhor || (ls.length === 1 ? ls[0] : null), novo = melhor && nota >= 0.5 ? melhor : null;
    if (velho) antes++; if (novo) agora++;
    if (velho && velho !== novo) { errados++; if (ex.length < 15) ex.push(r[iM] + ' ' + r[iN] + ' -> ' + idx[velho]); }
  });
  const msg = 'Itens com fotos de referência: antes ' + antes + ', agora ' + agora + ' | com fotos de outro modelo (corrigidos): ' + errados + '\n' + ex.join('\n');
  Logger.log(msg);
  return msg;
}
