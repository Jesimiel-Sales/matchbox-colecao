/**
 * Matchbox • Coleção Marco Geus
 * App web (Apps Script) cuja base de dados é a própria planilha.
 * A planilha é o espelho da aplicação: toda inclusão/edição/exclusão feita no app
 * é gravada nas abas Lesney e Superfast, com registro na aba _Log.
 */

const APP = {
  NOME: 'Matchbox • Coleção Marco Geus',
  HDR_ROW: 3,
  FIRST: 4,
  TECH_COL: 25, // coluna Y
  TZ: 'America/Sao_Paulo',
  BACKUP_DIAS: 30
};

// Mapa de colunas (1 = A). Campos técnicos ficam a partir da coluna Y em ambas as abas.
const SCHEMA = {
  Lesney: {
    prefix: 'L',
    codeField: 'mack',     // classificação Charlie Mack (modelo + variação)
    varField: 'mack2',     // classificação detalhada (quando houver)
    fields: [
      ['num', 1, 'Nº'], ['tipo', 2, 'Tipo'], ['mack', 3, 'Cl. Mack'], ['mack2', 4, 'Cl. Mack (detalhada)'],
      ['nome', 5, 'Nome'], ['cor', 6, 'Cor predominante'], ['base', 7, 'Base'], ['interior', 8, 'Interior'],
      ['vidro', 9, 'Vidro'], ['roda', 10, 'Roda'], ['outros', 11, 'Outros detalhes'], ['adesivo', 12, 'Adesivo/Decalque/Silks'],
      ['ano', 13, 'Ano'], ['dataAq', 14, 'Data de aquisição'], ['proced', 15, 'Procedência'], ['caixa', 16, 'Caixa (C.)'],
      ['valor', 17, 'Valor pago'], ['cotacao', 18, 'Cotação U$'], ['cons', 19, 'Conservação (Cl.)'], ['obs', 20, 'Obs.'],
      ['vendedor', 21, 'Adquirido de']
    ]
  },
  Superfast: {
    prefix: 'S',
    codeField: 'mack',
    varField: null,
    fields: [
      ['num', 1, 'Nº'], ['tipo', 2, 'Tipo'], ['mack', 4, 'C. C. Mack'], ['man', 5, 'MAN'],
      ['nome', 6, 'Nome'], ['cor', 7, 'Cor predominante'], ['base', 8, 'Base'], ['interior', 9, 'Interior'],
      ['vidro', 10, 'Vidros'], ['roda', 11, 'Rodas'], ['outros', 12, 'Outros detalhes'], ['adesivo', 13, 'Silks/Adesivo'],
      ['ano', 14, 'Ano'], ['dataAq', 15, 'Data de aquisição'], ['proced', 16, 'Procedência'], ['caixa', 17, 'Caixa (C.)'],
      ['valor', 18, 'Valor pago'], ['cotacao', 19, 'Cotação U$'], ['cons', 20, 'Conservação (Cl.)'], ['obs', 21, 'Obs.'],
      ['vendedor', 22, 'Adquirido de'], ['nota', 23, 'Nota']
    ]
  }
};
const TECH_FIELDS = [['id', 0, 'ID'], ['fotoRef', 1, 'Foto ref.'], ['fotos', 2, 'Fotos coleção'], ['atualizado', 3, 'Atualizado em'], ['manuais', 4, 'Alterado manualmente'], ['capa', 5, 'Foto principal']];
const NUMERIC = { ano: true, valor: true };
const DATE_FIELDS = { dataAq: true };

/* ---------------------------------------------------------------- web app */

function doGet(e) {
  if (e && e.parameter && e.parameter.diag) {
    return HtmlService.createHtmlOutput('<pre style="font:14px monospace;white-space:pre-wrap;padding:12px">' +
      diagnostico().replace(/</g, '&lt;') + '</pre>').setTitle('Diagnóstico Matchbox');
  }
  const out = HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP.NOME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
  const icone = PropertiesService.getScriptProperties().getProperty('ICONE_URL');
  if (icone) { try { out.setFaviconUrl(icone); } catch (err) { /* ícone opcional */ } }
  return out;
}

/** Rode pelo editor (Executar > diagnostico) ou abra a URL do app com ?diag=1 no final. */
function diagnostico() {
  const out = [];
  const ok = (t, f) => { try { out.push('OK   ' + t + ': ' + f()); } catch (err) { out.push('ERRO ' + t + ': ' + err.message); } };
  ok('Planilha vinculada', () => { const s = SpreadsheetApp.getActiveSpreadsheet(); if (!s) throw new Error('script não está vinculado a uma planilha (crie-o em Extensões > Apps Script dentro da planilha)'); return s.getName(); });
  Object.keys(SCHEMA).concat(['_Log', '_Lixeira', '_Ref', '_BAMCA', '_Config']).forEach(n => ok('Aba ' + n, () => {
    const sh = n === '_Ref' || n === '_BAMCA' ? refAba_(n) : ss_().getSheetByName(n); if (!sh) throw new Error('aba não encontrada');
    if (!SCHEMA[n]) return 'presente';
    const last = lastDataRow_(sh);
    const ids = last - APP.FIRST + 1;
    if (sh.getMaxColumns() < width_()) throw new Error('só tem ' + sh.getMaxColumns() + ' colunas; precisa de ' + width_());
    if (ids <= 0) throw new Error('nenhum ID na coluna Y — a planilha usada é a Matchbox_Colecao_Marco gerada pelo projeto?');
    return 'última linha com ID: ' + last + ' (' + sh.getRange(APP.HDR_ROW, APP.TECH_COL).getValue() + ')';
  }));
  ['Index'].forEach(f => ok('Arquivo HTML ' + f, () => HtmlService.createHtmlOutputFromFile(f).getContent().length + ' caracteres'));
  ok('Página montada', () => HtmlService.createTemplateFromFile('Index').evaluate().getContent().length + ' caracteres');
  ok('getData', () => { const t0 = Date.now(), s = getData(); const d = JSON.parse(s);
    return d.sheets.Lesney.length + ' Lesney, ' + d.sheets.Superfast.length + ' Superfast, ' + Math.round(s.length / 1024) + ' KB em ' + (Date.now() - t0) + ' ms'; });
  const txt = out.join('\n');
  console.log(txt);
  return txt;
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Matchbox App')
    .addItem('1. Instalar (pastas e cópia diária)', 'instalar')
    .addItem('2. Importar referência e fotos (uma vez, segundo plano)', 'iniciarReferenciaLocal')
    .addSeparator()
    .addItem('Fazer cópia de segurança agora', 'backupAgora')
    .addItem('Restaurar uma cópia de segurança…', 'restaurarPeloMenu')
    .addItem('Verificar integridade dos dados', 'verificarPeloMenu')
    .addItem('Gerar planilha da coleção (cópia limpa)', 'gerarPlanilhaPeloMenu')
    .addSeparator()
    .addItem('Configurar reconhecimento por foto (chave Gemini)…', 'configurarReconhecimento')
    .addItem('Publicar ícone do app', 'instalarIcone')
    .addItem('Atualizar dados (metadados) de todas as fotos', 'atualizarMetadadosFotos')
    .addToUi();
}

/* ---------------------------------------------------------------- leitura */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID')); }
const COLS_OK_ = {};
function sheet_(name) {
  if (!SCHEMA[name]) throw new Error('Aba desconhecida: ' + name);
  const sh = ss_().getSheetByName(name);
  if (!COLS_OK_[name]) { garantirColunas_(sh); COLS_OK_[name] = true; }
  return sh;
}
/** Garante as colunas técnicas (Y em diante), criando as que faltarem em planilhas de versões anteriores. */
function garantirColunas_(sh) {
  const w = APP.TECH_COL + TECH_FIELDS.length - 1, max = sh.getMaxColumns();
  if (max < w) sh.insertColumnsAfter(max, w - max);
  const hdr = sh.getRange(APP.HDR_ROW, APP.TECH_COL, 1, TECH_FIELDS.length);
  const v = hdr.getValues()[0];
  if (TECH_FIELDS.some((f, i) => v[i] !== f[2])) hdr.setValues([TECH_FIELDS.map(f => f[2])]);
}
function lastDataRow_(sh) {
  const last = sh.getLastRow();
  if (last < APP.FIRST) return APP.FIRST - 1;
  const ids = sh.getRange(APP.FIRST, APP.TECH_COL, last - APP.FIRST + 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0]) return APP.FIRST + i;
  return APP.FIRST - 1;
}
function width_() { return APP.TECH_COL + TECH_FIELDS.length - 1; }

function fmtCell_(key, v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, APP.TZ, key === 'atualizado' ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  return v;
}

function rowToItem_(name, row) {
  const s = SCHEMA[name], it = {};
  s.fields.forEach(f => it[f[0]] = fmtCell_(f[0], row[f[1] - 1]));
  TECH_FIELDS.forEach(f => it[f[0]] = fmtCell_(f[0], row[APP.TECH_COL - 1 + f[1]]));
  return it;
}

/** Carga inicial do app: esquema + todos os itens (em arrays para reduzir o tráfego).
 *  Devolve texto JSON: evita as limitações de serialização do google.script.run. */
function getData() {
  const out = { schema: {}, sheets: {}, geradoEm: new Date().toISOString() };
  Object.keys(SCHEMA).forEach(name => {
    const s = SCHEMA[name];
    const keys = s.fields.map(f => f[0]).concat(TECH_FIELDS.map(f => f[0]));
    out.schema[name] = { keys: keys, labels: s.fields.map(f => f[2]).concat(TECH_FIELDS.map(f => f[2])),
      codeField: s.codeField, varField: s.varField };
    const sh = sheet_(name), last = lastDataRow_(sh), rows = [];
    if (last >= APP.FIRST) {
      sh.getRange(APP.FIRST, 1, last - APP.FIRST + 1, width_()).getValues().forEach(r => {
        if (!r[APP.TECH_COL - 1]) return;
        const it = rowToItem_(name, r);
        rows.push(keys.map(k => it[k]));
      });
    }
    out.sheets[name] = rows;
  });
  return JSON.stringify(out);
}

/* ---------------------------------------------------------------- escrita */

function findRow_(sh, id) {
  const last = lastDataRow_(sh);
  if (last < APP.FIRST) return 0;
  const cell = sh.getRange(APP.FIRST, APP.TECH_COL, last - APP.FIRST + 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}

function nextId_(name, sh) {
  const last = lastDataRow_(sh), p = SCHEMA[name].prefix + '-';
  let max = 0;
  if (last >= APP.FIRST) sh.getRange(APP.FIRST, APP.TECH_COL, last - APP.FIRST + 1, 1).getValues()
    .forEach(r => { const m = String(r[0]).match(/-(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  return p + ('0000' + (max + 1)).slice(-5);
}

function numKey_(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

/** Posição de inserção: logo após o último item com o mesmo Nº (mantém a ordem do Marco). */
function insertPosition_(sh, item) {
  const last = lastDataRow_(sh);
  if (last < APP.FIRST) return APP.FIRST - 1;
  const nums = sh.getRange(APP.FIRST, 1, last - APP.FIRST + 1, 1).getValues().map(r => r[0]);
  const target = String(item.num).trim(), tNum = numKey_(item.num);
  let pos = -1, lastSmaller = -1;
  let current = null; // Nº vigente (linhas sem Nº herdam o de cima)
  nums.forEach((v, i) => {
    if (v !== '' && v !== null) current = v;
    if (current !== null && String(current).trim() === target) pos = i;
    const c = numKey_(current);
    if (tNum !== null && c !== null && c < tNum) lastSmaller = i;
  });
  if (pos >= 0) return APP.FIRST + pos;
  if (lastSmaller >= 0) return APP.FIRST + lastSmaller;
  return last;
}

function cleanValue_(key, v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') v = v.replace(/\s+/g, ' ').trim();
  if (v === '') return '';
  if (key === 'cons') return String(v).toUpperCase();
  if (key === 'num') { const n = Number(v); return isNaN(n) ? v : n; }
  if (NUMERIC[key]) { const n = Number(String(v).replace(',', '.')); return isNaN(n) ? v : n; }
  if (DATE_FIELDS[key]) {
    const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3]) : v;
  }
  return v;
}

function writeItem_(name, sh, row, item) {
  const s = SCHEMA[name];
  const range = sh.getRange(row, 1, 1, width_());
  const vals = range.getValues()[0];
  s.fields.forEach(f => { if (f[0] in item) vals[f[1] - 1] = cleanValue_(f[0], item[f[0]]); });
  vals[APP.TECH_COL - 1] = item.id;
  if ('fotoRef' in item) vals[APP.TECH_COL] = item.fotoRef || '';
  if ('fotos' in item) vals[APP.TECH_COL + 1] = item.fotos || '';
  vals[APP.TECH_COL + 2] = new Date();
  if ('manuais' in item) vals[APP.TECH_COL + 3] = item.manuais || '';
  range.setValues([vals]);
  return rowToItem_(name, vals);
}

function log_(acao, name, before, after) {
  const it = after || before || {};
  ss_().getSheetByName('_Log').appendRow([new Date(), acao, name, it.id || '', it.mack || '', it.nome || '',
    before ? JSON.stringify(before) : '', after ? JSON.stringify(after) : '']);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('A planilha está ocupada. Tente de novo em alguns segundos.');
  try { return fn(); } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}

/** Inclui (sem id) ou edita (com id) um item. Retorna o item como ficou gravado. */
function saveItem(name, item) {
  return withLock_(() => {
    const sh = sheet_(name);
    if (!item.nome || !String(item.nome).trim()) throw new Error('Informe o nome do modelo.');
    if (item.id) {
      const row = findRow_(sh, item.id);
      if (!row) throw new Error('Item ' + item.id + ' não encontrado na planilha (pode ter sido excluído).');
      const before = rowToItem_(name, sh.getRange(row, 1, 1, width_()).getValues()[0]);
      const after = writeItem_(name, sh, row, item);
      log_('editar', name, before, after);
      atualizarDescricoesFotos_(name, after);
      return after;
    }
    item.id = nextId_(name, sh);
    const pos = insertPosition_(sh, item);
    sh.insertRowAfter(pos);
    const after = writeItem_(name, sh, pos + 1, item);
    log_('incluir', name, null, after);
    return after;
  });
}

/** Exclusão segura: a linha vai para a aba _Lixeira e pode ser restaurada. */
function deleteItem(name, id) {
  return withLock_(() => {
    const sh = sheet_(name), row = findRow_(sh, id);
    if (!row) throw new Error('Item não encontrado: ' + id);
    const item = rowToItem_(name, sh.getRange(row, 1, 1, width_()).getValues()[0]);
    ss_().getSheetByName('_Lixeira').appendRow([new Date(), name, id, row, JSON.stringify(item)]);
    sh.deleteRow(row);
    log_('excluir', name, item, null);
    return id;
  });
}

function listTrash() {
  const sh = ss_().getSheetByName('_Lixeira'), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues().map((r, i) => ({
    linha: i + 2, quando: fmtCell_('atualizado', r[0]), aba: r[1], id: r[2], item: JSON.parse(r[4] || '{}')
  })).reverse();
}

function restoreItem(linhaLixeira) {
  return withLock_(() => {
    const lx = ss_().getSheetByName('_Lixeira');
    const r = lx.getRange(linhaLixeira, 1, 1, 5).getValues()[0];
    const name = r[1], item = JSON.parse(r[4]);
    const sh = sheet_(name);
    if (findRow_(sh, item.id)) throw new Error('Este item já está na coleção.');
    const pos = insertPosition_(sh, item);
    sh.insertRowAfter(pos);
    const after = writeItem_(name, sh, pos + 1, item);
    lx.deleteRow(linhaLixeira);
    log_('restaurar', name, null, after);
    return { aba: name, item: after };
  });
}

/** Grava a foto de referência escolhida. Não acessa sites: usa a cópia que já está no banco de imagens. */
function setFotoRef(name, id, url, driveId, daVariacao) {
  const valor = !url ? '' : (daVariacao ? 'var:' : '') + (driveId ? 'drive:' + driveId + '|' + url : url);
  return withLock_(() => {
    const sh = sheet_(name), row = findRow_(sh, id);
    if (!row) throw new Error('Item não encontrado: ' + id);
    sh.getRange(row, APP.TECH_COL + 1).setValue(valor);
    return valor;
  });
}

/* ---------------------------------------------------------------- fotos da coleção (Drive) */

function cfgGet_(k) {
  const sh = ss_().getSheetByName('_Config'), vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const r = vals.find(v => v[0] === k);
  return r ? r[1] : '';
}
function cfgSet_(k, v) {
  const sh = ss_().getSheetByName('_Config'), vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  const i = vals.findIndex(x => x[0] === k);
  if (i >= 0) sh.getRange(i + 1, 2).setValue(v); else sh.appendRow([k, v]);
}
function folder_(key, nome) {
  const id = cfgGet_(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recria */ } }
  const parent = DriveApp.getFileById(ss_().getId()).getParents();
  const base = parent.hasNext() ? parent.next() : DriveApp.getRootFolder();
  const f = base.createFolder(nome);
  cfgSet_(key, f.getId());
  return f;
}

/** Recebe a foto já reduzida no celular (data URL JPEG) e anexa ao item. */
function uploadPhoto(name, id, dataUrl) {
  const m = String(dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) throw new Error('Formato de imagem inválido.');
  const folder = folder_('pasta_fotos_id', 'Matchbox • Fotos da coleção');
  const stamp = Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd-HHmmss');
  const sh0 = sheet_(name), row0 = findRow_(sh0, id);
  const item0 = row0 ? rowToItem_(name, sh0.getRange(row0, 1, 1, width_()).getValues()[0]) : { id: id };
  let bytes = Utilities.base64Decode(m[2]);
  try { bytes = comExif_(bytes, infoFoto_(name, item0)); } catch (e) { /* sem EXIF: segue com a foto original */ }
  const blob = Utilities.newBlob(bytes, 'image/jpeg', nomeFoto_(item0, stamp));
  const file = folder.createFile(blob);
  try { file.setDescription(descricaoFoto_(name, item0)); } catch (e) { /* opcional */ }
  return withLock_(() => {
    const sh = sheet_(name), row = findRow_(sh, id);
    if (!row) { file.setTrashed(true); throw new Error('Item não encontrado: ' + id); }
    const cell = sh.getRange(row, APP.TECH_COL + 2);
    const list = String(cell.getValue() || '').split(';').filter(String);
    list.push(file.getId());
    cell.setValue(list.join(';'));
    log_('foto+', name, null, { id: id, foto: file.getId() });
    return list.join(';');
  });
}

function removePhoto(name, id, fileId) {
  return withLock_(() => {
    const sh = sheet_(name), row = findRow_(sh, id);
    if (!row) throw new Error('Item não encontrado: ' + id);
    const cell = sh.getRange(row, APP.TECH_COL + 2);
    const list = String(cell.getValue() || '').split(';').filter(x => x && x !== fileId);
    cell.setValue(list.join(';'));
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* já removida */ }
    log_('foto-', name, { id: id, foto: fileId }, null);
    return list.join(';');
  });
}

/** Miniatura de uma foto do Drive como data URL (funciona mesmo sem compartilhamento público). */
function getDrivePhoto(fileId, size) {
  const cache = CacheService.getScriptCache(), key = 'dp:' + fileId + ':' + (size || 600);
  const hit = cache.get(key);
  if (hit) return hit;
  // as fotos já chegam reduzidas pelo celular (lado maior 1200 px), então o arquivo original serve
  const blob = DriveApp.getFileById(fileId).getBlob();
  const data = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  if (data.length < 95000) cache.put(key, data, 21600);
  return data;
}

/* ---------------------------------------------------------------- instalação */

function instalar() {
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss_().getId());
  folder_('pasta_fotos_id', 'Matchbox • Fotos da coleção');
  folder_('pasta_backup_id', 'Matchbox • Backups');
  try { instalarIcone(true); } catch (err) { /* ícone é opcional */ }
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backupDiario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupDiario').timeBased().everyDays(1).atHour(3).inTimezone(APP.TZ).create();
  PropertiesService.getScriptProperties().setProperty('gatilho_backup', new Date().toISOString());
  backupDiario();
  try { instalarVerificacaoSemanal(); } catch (err) { /* opcional */ }
  avisar_('Instalação concluída: pastas criadas, cópia diária às 3h ativada e a cópia de hoje já foi feita.');
}

function avisar_(msg) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* sem interface (gatilho ou app) */ } }

/* ---------------------------------------------------------------- cópias de segurança (rodízio de 7 dias) */

const DIAS_SEMANA = ['', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'];
const PREFIXO_COPIA = 'Cópia Matchbox — ';

/** Uma cópia completa da planilha por dia da semana. A cópia de cada dia substitui a da semana anterior
 *  (a substituída vai para a lixeira do Drive, onde fica mais 30 dias). */
function backupDiario(manual) {
  const folder = folder_('pasta_backup_id', 'Matchbox • Backups');
  const agora = new Date();
  const nome = PREFIXO_COPIA + (manual === true ? 'Manual' : DIAS_SEMANA[+Utilities.formatDate(agora, APP.TZ, 'u')]);
  const it = folder.getFilesByName(nome);
  while (it.hasNext()) { const f = it.next(); try { f.setTrashed(true); } catch (e) { /* cópia de outro dono */ } }
  const copia = DriveApp.getFileById(ss_().getId()).makeCopy(nome, folder);
  copia.setDescription('Cópia feita em ' + Utilities.formatDate(agora, APP.TZ, 'dd/MM/yyyy HH:mm'));
  PropertiesService.getScriptProperties().setProperty('ultima_copia', agora.toISOString());
  return nome;
}

function backupAgora() {
  const nome = backupDiario(true);
  avisar_('Cópia de segurança feita: "' + nome + '" na pasta Matchbox • Backups.');
  return nome;
}

function listarBackups_() {
  const folder = folder_('pasta_backup_id', 'Matchbox • Backups'), out = [], it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    out.push({ id: f.getId(), nome: f.getName(), url: f.getUrl(), data: f.getDateCreated().getTime(),
      quando: Utilities.formatDate(f.getDateCreated(), APP.TZ, 'dd/MM/yyyy HH:mm'),
      atual: f.getName().indexOf(PREFIXO_COPIA) === 0 });
  }
  return out.sort((a, b) => b.data - a.data);
}

/** Restaura as abas Lesney e Superfast a partir de uma cópia. Antes, guarda o estado atual
 *  numa cópia "Antes da restauração", para que a própria restauração possa ser desfeita. */
function restaurarBackup(fileId) {
  const origem = SpreadsheetApp.openById(fileId);
  const nomeOrigem = DriveApp.getFileById(fileId).getName();
  const folder = folder_('pasta_backup_id', 'Matchbox • Backups');
  const nomeSeg = PREFIXO_COPIA + 'Antes da restauração';
  const it = folder.getFilesByName(nomeSeg);
  while (it.hasNext()) { const f = it.next(); try { f.setTrashed(true); } catch (e) { /* cópia de outro dono */ } }
  DriveApp.getFileById(ss_().getId()).makeCopy(nomeSeg, folder).setDescription('Estado da coleção antes de restaurar "' + nomeOrigem + '"');
  return withLock_(() => {
    const res = {};
    Object.keys(SCHEMA).forEach(name => {
      const src = origem.getSheetByName(name);
      if (!src) throw new Error('A cópia não tem a aba ' + name + '.');
      const dst = sheet_(name), n = Math.max(src.getLastRow(), 1), w = Math.min(width_(), src.getMaxColumns());
      const vals = src.getRange(1, 1, n, w).getValues();
      const dLast = Math.max(dst.getLastRow(), 1);
      dst.getRange(1, 1, dLast, w).clearContent();
      if (dst.getMaxRows() < n) dst.insertRowsAfter(dst.getMaxRows(), n - dst.getMaxRows());
      dst.getRange(1, 1, n, w).setValues(vals);
      res[name] = vals.slice(APP.FIRST - 1).filter(r => r[APP.TECH_COL - 1]).length;
    });
    log_('restaurar cópia', 'Lesney/Superfast', null, { id: '', nome: nomeOrigem, mack: '' });
    return { origem: nomeOrigem, itens: res, seguranca: nomeSeg };
  });
}

/** Mesma restauração, pelo menu da planilha (serve mesmo se o app não abrir). */
function restaurarPeloMenu() {
  const ui = SpreadsheetApp.getUi(), lista = listarBackups_();
  if (!lista.length) { ui.alert('Nenhuma cópia encontrada na pasta Matchbox • Backups.'); return; }
  const txt = lista.map((b, i) => (i + 1) + '. ' + b.nome + '  (' + b.quando + ')').join('\n');
  const r = ui.prompt('Restaurar cópia de segurança', 'Digite o número da cópia que deseja restaurar:\n\n' + txt, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const b = lista[parseInt(r.getResponseText(), 10) - 1];
  if (!b) { ui.alert('Número inválido.'); return; }
  if (ui.alert('Confirmar', 'As abas Lesney e Superfast vão voltar a ser como estavam em "' + b.nome + '" (' + b.quando + ').\n\nO estado atual será guardado antes, na cópia "Antes da restauração". Continuar?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const res = restaurarBackup(b.id);
  ui.alert('Restauração concluída: ' + res.itens.Lesney + ' itens Lesney e ' + res.itens.Superfast + ' Superfast.');
}

/* ---------------------------------------------------------------- integridade */

/** Confere IDs: linhas preenchidas sem ID (ex.: digitadas direto na planilha) e IDs repetidos.
 *  Com corrigir=true, dá ID às linhas sem ID e renumera as repetidas. */
function verificarIntegridade(corrigir) {
  const run = () => {
    const out = { abas: {}, problemas: 0, corrigidos: 0 };
    Object.keys(SCHEMA).forEach(name => {
      const s = SCHEMA[name], sh = sheet_(name), last = sh.getLastRow();
      const r = { itens: 0, semId: [], duplicados: [] };
      if (last >= APP.FIRST) {
        const vals = sh.getRange(APP.FIRST, 1, last - APP.FIRST + 1, width_()).getValues();
        const nomeCol = s.fields.find(f => f[0] === 'nome')[1] - 1, codCol = s.fields.find(f => f[0] === 'mack')[1] - 1;
        const vistos = {};
        let max = 0;
        vals.forEach(v => { const m = String(v[APP.TECH_COL - 1]).match(/-(\d+)$/); if (m) max = Math.max(max, +m[1]); });
        vals.forEach((v, i) => {
          const id = v[APP.TECH_COL - 1], temDado = v[nomeCol] !== '' || v[codCol] !== '';
          if (!temDado) return;
          r.itens++;
          let precisa = false;
          if (!id) { r.semId.push(APP.FIRST + i); precisa = true; }
          else if (vistos[id]) { r.duplicados.push(id + ' (linha ' + (APP.FIRST + i) + ')'); precisa = true; }
          else vistos[id] = true;
          if (precisa && corrigir) {
            const novo = s.prefix + '-' + ('0000' + (++max)).slice(-5);
            sh.getRange(APP.FIRST + i, APP.TECH_COL).setValue(novo);
            vistos[novo] = true; out.corrigidos++;
          }
        });
      }
      out.problemas += r.semId.length + r.duplicados.length;
      out.abas[name] = r;
    });
    if (corrigir && out.corrigidos) log_('integridade', 'Lesney/Superfast', null, { id: '', nome: out.corrigidos + ' IDs atribuídos', mack: '' });
    return out;
  };
  return corrigir ? withLock_(run) : run();
}

function verificarPeloMenu() {
  const ui = SpreadsheetApp.getUi(), v = verificarIntegridade(false);
  const resumo = Object.keys(v.abas).map(n => n + ': ' + v.abas[n].itens + ' itens, ' + v.abas[n].semId.length + ' sem ID, ' + v.abas[n].duplicados.length + ' IDs repetidos').join('\n');
  if (!v.problemas) { ui.alert('Tudo certo.\n\n' + resumo); return; }
  if (ui.alert('Foram encontrados ' + v.problemas + ' problema(s).\n\n' + resumo + '\n\nCorrigir agora (dar ID às linhas sem ID)?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    const c = verificarIntegridade(true);
    ui.alert(c.corrigidos + ' linha(s) corrigida(s).');
  }
}

/* ---------------------------------------------------------------- painel "Segurança dos dados" */

function getSeguranca() {
  const props = PropertiesService.getScriptProperties();
  let backups = [], pastaUrl = '';
  try { backups = listarBackups_(); pastaUrl = folder_('pasta_backup_id', 'Matchbox • Backups').getUrl(); } catch (e) {}
  const atuais = backups.filter(b => b.atual && !/Antes da restaura/.test(b.nome));
  const ultima = atuais.length ? atuais[0] : null;
  const horas = ultima ? (Date.now() - ultima.data) / 36e5 : null;
  const integ = verificarIntegridade(false);
  return JSON.stringify({
    backups: backups, ultima: ultima, horasDesdeUltima: horas,
    gatilhoAtivo: !!props.getProperty('gatilho_backup') || ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupDiario'),
    integridade: integ, pastaUrl: pastaUrl, planilhaUrl: ss_().getUrl(),
    semanal: (function () { try { return JSON.parse(cfgGet_('verificacao_semanal') || 'null'); } catch (e) { return null; } })(),
    referencia: JSON.parse(props.getProperty('ref_local') || 'null'),
    referenciaRodando: (function () { const r = JSON.parse(props.getProperty('ref_local') || 'null'); return !!r && r.fase !== 'concluída'; })()
  });
}

function restaurarBackupApp(fileId) { return JSON.stringify(restaurarBackup(fileId)); }
function corrigirIntegridadeApp() { return JSON.stringify(verificarIntegridade(true)); }


/* ---------------------------------------------------------------- ícone do app */

/** Publica o ícone (PNG embutido em ICONE_PNG_B64) no Drive, com link público, e usa como favicon do app. */
function instalarIcone(silencioso) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('ICONE_ID'), file = null;
  if (id) { try { file = DriveApp.getFileById(id); if (file.isTrashed()) file = null; } catch (e) { file = null; } }
  if (!file) {
    const blob = Utilities.newBlob(Utilities.base64Decode(ICONE_PNG_B64), 'image/png', 'Matchbox App - icone.png');
    file = folder_('pasta_fotos_id', 'Matchbox • Fotos da coleção').createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    id = file.getId();
  }
  const url = 'https://lh3.googleusercontent.com/d/' + id + '=s192#icone.png';
  props.setProperty('ICONE_ID', id); props.setProperty('ICONE_URL', url);
  if (!silencioso) avisar_('Ícone publicado. Ele aparece na aba do navegador e ao adicionar o app à tela inicial.');
  return url;
}

/* ---------------------------------------------------------------- reconhecimento por foto (Gemini) */

const IA = { MODELO: 'gemini-3.8-flash', RESERVA: ['gemini-3.5-flash', 'gemini-flash-latest', 'gemini-2.5-flash'], URL: 'https://generativelanguage.googleapis.com/v1beta/models/' };
function iaCfg_() {
  const p = PropertiesService.getScriptProperties();
  return { chave: p.getProperty('GEMINI_API_KEY'), modelo: p.getProperty('GEMINI_MODELO') || IA.MODELO };
}

/** Menu da planilha: guarda a chave do Google AI Studio nas propriedades do script (não fica na planilha). */
function configurarReconhecimento() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Reconhecimento por foto',
    'Cole a chave de API do Google AI Studio (aistudio.google.com > Get API key).\nDeixe em branco e clique OK para desligar o recurso.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const chave = r.getResponseText().trim(), p = PropertiesService.getScriptProperties();
  if (!chave) { p.deleteProperty('GEMINI_API_KEY'); ui.alert('Reconhecimento por foto desligado.'); return; }
  const resp = UrlFetchApp.fetch(IA.URL + encodeURIComponent(iaCfg_().modelo), { headers: { 'x-goog-api-key': chave }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) { ui.alert('A chave não funcionou (código ' + resp.getResponseCode() + '). Nada foi guardado.\n\n' + resp.getContentText().slice(0, 300)); return; }
  p.setProperty('GEMINI_API_KEY', chave);
  ui.alert('Pronto. O botão da câmera na busca do app já reconhece miniaturas por foto.');
}

function getIconeData() { return 'data:image/png;base64,' + ICONE_PNG_B64; }

function iaDisponivel() { return !!iaCfg_().chave; }

const PROMPT_IA_ = [
  'You are an expert on Lesney-era Matchbox 1-75 diecast models (regular wheels 1953-1969 and Superfast 1969-1982).',
  'Identify the toy in the photo. Reply ONLY with JSON in this shape:',
  '{"e_miniatura":true,"candidatos":[{"numero":1,"nome":"Road Roller","serie":"Lesney|Superfast|outra","ano":1958,',
  '"cor_pt":"verde escuro","cor_en":"dark green","rodas":"metal|plastico cinza|plastico preto|superfast|outra",',
  '"texto_base":"text readable on the baseplate or empty","confianca":0.8}],"observacao_pt":"short note in Portuguese"}',
  'Give up to 3 candidates ordered by likelihood. "numero" is the 1-75 series number. Use the English model name as on the box.',
  'If the photo shows a box, use the box text. If it is not a diecast toy, set e_miniatura to false and candidatos to [].'
].join('\n');

/** Recebe uma foto (data URL, já reduzida pelo celular) e devolve JSON com os candidatos. */
function reconhecerImagem(dataUrl) {
  const cfg = iaCfg_();
  if (!cfg.chave) return JSON.stringify({ ok: false, motivo: 'sem_chave' });
  const m = String(dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) throw new Error('Formato de imagem inválido.');
  const corpo = {
    contents: [{ role: 'user', parts: [{ text: PROMPT_IA_ }, { inlineData: { mimeType: m[1], data: m[2] } }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.2 }
  };
  // Tenta o modelo configurado e, se estiver sobrecarregado (503/429/500) ou indisponível, outros modelos rápidos.
  const modelos = [cfg.modelo].concat(IA.RESERVA.filter(x => x !== cfg.modelo)).slice(0, 4);
  let code = 0, txt = '', usado = '';
  for (let i = 0; i < modelos.length; i++) {
    usado = modelos[i];
    const resp = UrlFetchApp.fetch(IA.URL + encodeURIComponent(usado) + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': cfg.chave },
      payload: JSON.stringify(corpo), muteHttpExceptions: true
    });
    code = resp.getResponseCode(); txt = resp.getContentText();
    if (code === 200) break;
    if ([429, 500, 503, 504, 404].indexOf(code) < 0) break; // erro de chave/pedido: não adianta trocar
  }
  if (code !== 200) {
    let msg = ''; try { msg = JSON.parse(txt).error.message; } catch (e) { msg = txt.slice(0, 200); }
    if (code === 503 || code === 429) msg = 'o serviço do Google está sobrecarregado agora; tente de novo em alguns minutos.';
    return JSON.stringify({ ok: false, motivo: 'erro', mensagem: 'Reconhecimento indisponível (' + code + '): ' + msg });
  }
  const d = JSON.parse(txt), partes = (((d.candidates || [])[0] || {}).content || {}).parts || [];
  const bruto = partes.map(p => p.text || '').join('').replace(/^```(json)?|```$/g, '').trim();
  let r; try { r = JSON.parse(bruto); } catch (e) { return JSON.stringify({ ok: false, motivo: 'erro', mensagem: 'Resposta não reconhecida.' }); }
  return JSON.stringify({ ok: true, r: r });
}

const ICONE_PNG_B64 = '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';


/* ---------------------------------------------------------------- metadados das fotos da coleção */

/** Código Charlie Mack do item: na Lesney é a classificação detalhada (mack2); a coluna "Cl." fica como referência. */
function codigoMack_(name, item) { return name === 'Lesney' ? (item.mack2 || item.mack || '') : (item.mack || ''); }

function linhasFoto_(name, it) {
  const cod = codigoMack_(name, it);
  const l = [
    'Matchbox ' + (name === 'Lesney' ? 'Lesney 1-75' : 'Superfast') + ' No. ' + it.num + ' - ' + it.nome,
    'Código Charlie Mack: ' + cod + (name === 'Lesney' && it.mack && it.mack !== cod ? ' (Cl. ' + it.mack + ')' : '') + (it.man ? ' - MAN ' + it.man : ''),
    ['Cor', it.cor], ['Rodas', it.roda], ['Base', it.base], ['Interior', it.interior], ['Vidro', it.vidro],
    ['Outros detalhes', it.outros], ['Adesivo', it.adesivo], ['Ano', it.ano], ['Procedência', it.proced],
    ['Conservação', it.cons], ['Caixa', it.caixa], ['Tipo', it.tipo], ['Obs.', it.obs],
    'ID ' + it.id + ' - Coleção Marco Geus'
  ];
  return l.map(x => Array.isArray(x) ? (x[1] !== '' && x[1] != null ? x[0] + ': ' + x[1] : '') : x).filter(String);
}
function descricaoFoto_(name, it) { return linhasFoto_(name, it).join('\n'); }
function infoFoto_(name, it) {
  const l = linhasFoto_(name, it);
  return { titulo: l[0], descricao: l.join(' | '), comentario: l.join('\n'), autor: 'Colecao Marco Geus' };
}
function nomeFoto_(it, stamp) {
  const limpo = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40);
  return [it.id, limpo(codigoMack_(it.id && it.id[0] === 'S' ? 'Superfast' : 'Lesney', it)), limpo(it.nome), stamp].filter(String).join('_') + '.jpg';
}

/** Atualiza descrição e nome no Drive das fotos de um item (chamado ao salvar). */
function atualizarDescricoesFotos_(name, it) {
  String(it.fotos || '').split(';').filter(String).forEach(fid => {
    try {
      const f = DriveApp.getFileById(fid);
      f.setDescription(descricaoFoto_(name, it));
      const m = f.getName().match(/(\d{8}-\d{6})/);
      if (m) f.setName(nomeFoto_(it, m[1]));
    } catch (e) { /* foto apagada ou sem permissão: ignora */ }
  });
}

/** Menu/editor: grava os dados do cadastro em todas as fotos da coleção já enviadas. */
function atualizarMetadadosFotos() {
  let n = 0;
  Object.keys(SCHEMA).forEach(name => {
    const sh = sheet_(name), last = lastDataRow_(sh);
    if (last < APP.FIRST) return;
    sh.getRange(APP.FIRST, 1, last - APP.FIRST + 1, width_()).getValues().forEach(r => {
      const it = rowToItem_(name, r);
      if (it.id && it.fotos) { atualizarDescricoesFotos_(name, it); n += String(it.fotos).split(';').filter(String).length; }
    });
  });
  avisar_(n + ' foto(s) atualizada(s) com os dados do cadastro.');
  return n;
}

/**
 * Insere um bloco EXIF (APP1) no JPEG com os dados do item:
 * ImageDescription e Artist (ASCII), XPTitle e XPComment (Windows, com acentos) e UserComment (Unicode).
 */
function comExif_(signedBytes, info) {
  const b = signedBytes.map(x => x & 255);
  if (b[0] !== 0xFF || b[1] !== 0xD8) return signedBytes;
  const ascii = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\x20-\x7e\n]/g, '?').slice(0, 1800)
    .split('').map(c => c.charCodeAt(0)).concat([0]);
  const ucs = (s, le) => { const o = []; String(s || '').slice(0, 1800).split('').forEach(ch => { const c = ch.charCodeAt(0); if (le) o.push(c & 255, c >> 8); else o.push(c >> 8, c & 255); }); return o; };
  const e0 = [[0x010E, 2, ascii(info.descricao)], [0x013B, 2, ascii(info.autor)], [0x8769, 4, null],
    [0x9C9B, 1, ucs(info.titulo, true).concat([0, 0])], [0x9C9C, 1, ucs(info.comentario, true).concat([0, 0])]];
  const eE = [[0x9286, 7, [85, 78, 73, 67, 79, 68, 69, 0].concat(ucs(info.comentario, false))]];
  const t = [];
  const u16 = (a, v) => a.push((v >> 8) & 255, v & 255);
  const u32 = (a, v) => a.push((v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255);
  const tam = es => es.reduce((s, e) => s + (e[2] && e[2].length > 4 ? e[2].length + (e[2].length & 1) : 0), 0);
  const off0 = 8, dados0 = off0 + 2 + 12 * e0.length + 4, offE = dados0 + tam(e0), dadosE = offE + 2 + 12 * eE.length + 4;
  const escreverIfd = (es, inicioDados, prox) => {
    const ifd = [], dados = []; let pos = inicioDados;
    u16(ifd, es.length);
    es.forEach(e => {
      u16(ifd, e[0]); u16(ifd, e[1]);
      if (e[2] === null) { u32(ifd, 1); u32(ifd, offE); return; }
      u32(ifd, e[2].length);
      if (e[2].length <= 4) { const v = e[2].concat([0, 0, 0, 0]).slice(0, 4); v.forEach(x => ifd.push(x)); }
      else { u32(ifd, pos); e[2].forEach(x => dados.push(x)); if (e[2].length & 1) dados.push(0); pos += e[2].length + (e[2].length & 1); }
    });
    u32(ifd, prox);
    return ifd.concat(dados);
  };
  t.push(0x4D, 0x4D, 0, 0x2A); u32(t, off0);
  const tiff = t.concat(escreverIfd(e0, dados0, 0), escreverIfd(eE, dadosE, 0));
  const app1 = [0xFF, 0xE1]; u16(app1, tiff.length + 8);
  if (tiff.length + 8 > 65000) return signedBytes;
  const seg = app1.concat([0x45, 0x78, 0x69, 0x66, 0, 0], tiff);
  let ins = 2;
  if (b[2] === 0xFF && b[3] === 0xE0) ins = 4 + ((b[4] << 8) | b[5]);
  const out = b.slice(0, ins).concat(seg, b.slice(ins));
  return out.map(x => x > 127 ? x - 256 : x);
}

/** Uma vez: itens Lesney cujo modelo pela Cl. difere do código Mack ficaram com foto de referência do modelo errado. */
function corrigirReferenciasMack() {
  const sh = sheet_('Lesney'), last = lastDataRow_(sh); let n = 0;
  if (last < APP.FIRST) return 0;
  const vals = sh.getRange(APP.FIRST, 1, last - APP.FIRST + 1, width_()).getValues();
  const p = s => { const m = String(s || '').trim().match(/^0*(\d{1,3})\s*([A-Za-z])/); return m ? m[1] + m[2].toUpperCase() : ''; };
  vals.forEach((r, i) => {
    const it = rowToItem_('Lesney', r);
    if (it.mack2 && p(it.mack) && p(it.mack2) && p(it.mack) !== p(it.mack2) && !/^var:/.test(it.fotoRef || '')) {
      sh.getRange(APP.FIRST + i, APP.TECH_COL + 1).setValue(''); n++;
    }
  });
  return n;
}


/* ---------------------------------------------------------------- foto principal e enquadramento */

/** Grava a foto principal do item e o enquadramento da miniatura: "<id da foto ou ref>|x,y,largura,altura" (frações da imagem). */
function setCapa(name, id, valor) {
  return withLock_(() => {
    const sh = sheet_(name), row = findRow_(sh, id);
    if (!row) throw new Error('Item não encontrado: ' + id);
    const v = String(valor || '').slice(0, 200);
    sh.getRange(row, APP.TECH_COL + 5).setValue(v);
    return v;
  });
}

/** Editor: mostra no registro de execução se o reconhecimento por foto está configurado e responde. */
function diagnosticoIA() {
  const cfg = iaCfg_();
  Logger.log('Chave configurada: ' + (cfg.chave ? 'sim' : 'NÃO') + ' | modelo: ' + cfg.modelo);
  if (!cfg.chave) return 'sem chave';
  const r = UrlFetchApp.fetch(IA.URL.replace(/models\/$/, 'models?pageSize=200'), { headers: { 'x-goog-api-key': cfg.chave }, muteHttpExceptions: true });
  const lista = (JSON.parse(r.getContentText()).models || []).filter(x => (x.supportedGenerationMethods || []).indexOf('generateContent') >= 0).map(x => x.name.replace('models/', ''));
  Logger.log('Modelos com generateContent: ' + lista.filter(n => /flash|pro/.test(n)).join(', '));
  return r.getResponseCode();
}

/** Testa o reconhecimento com a foto mais recente da pasta de fotos (rodar pelo editor). */
function testarReconhecimento() {
  const pasta = folder_('pasta_fotos_id', 'Matchbox • Fotos da coleção');
  const it = pasta.getFiles(); let f = null;
  while (it.hasNext()) { const x = it.next(); if (!f || x.getDateCreated() > f.getDateCreated()) f = x; }
  if (!f) { Logger.log('sem fotos'); return; }
  const blob = f.getBlob();
  const d = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  Logger.log('Foto: ' + f.getName() + ' (' + Math.round(blob.getBytes().length / 1024) + ' KB)');
  Logger.log(reconhecerImagem(d).slice(0, 900));
}


/* ---------------------------------------------------------------- estabilidade (v1.11) */

/** Referência (Charlie Mack e BAMCA) fica numa planilha separada, para a planilha do Marco ficar leve.
 *  Enquanto a separação não for feita, tudo continua sendo lido da própria planilha. */
const REF_ABAS = ['_Ref', '_BAMCA'];
let REF_SS_ = null;
function refSs_() {
  if (REF_SS_) return REF_SS_;
  const id = cfgGet_('ref_planilha_id');
  if (id) { try { REF_SS_ = SpreadsheetApp.openById(id); return REF_SS_; } catch (e) { /* sem acesso: usa a principal */ } }
  REF_SS_ = ss_();
  return REF_SS_;
}
function refAba_(nome) { return refSs_().getSheetByName(nome) || ss_().getSheetByName(nome); }

/** Rodar uma vez pelo editor. Copia _Ref e _BAMCA para "Matchbox • Referência", confere linha a linha
 *  e só então apaga as abas da planilha principal. */
function separarReferencia() {
  const ss = ss_();
  let dest = null, id = cfgGet_('ref_planilha_id');
  if (id) { try { dest = SpreadsheetApp.openById(id); } catch (e) { dest = null; } }
  if (!dest) {
    dest = SpreadsheetApp.create('Matchbox • Referência (Charlie Mack e BAMCA)');
    const pais = DriveApp.getFileById(ss.getId()).getParents();
    if (pais.hasNext()) DriveApp.getFileById(dest.getId()).moveTo(pais.next());
    cfgSet_('ref_planilha_id', dest.getId());
  }
  dest.setSpreadsheetTimeZone(ss.getSpreadsheetTimeZone());
  let sobre = dest.getSheetByName('Sobre');
  if (!sobre) {
    sobre = dest.insertSheet('Sobre', 0);
    sobre.getRange(1, 1, 4, 1).setValues([['Referência do app Matchbox (não editar à mão).'], ['_Ref: lista do Charlie Mack e páginas de referência.'],
      ['_BAMCA: variações do BAMCA (bamca.org), cópia autorizada para uso privado.'], ['Ligada à planilha ' + ss.getName() + '.']]);
  }
  sobre.showSheet(); dest.setActiveSheet(sobre);
  const igual = (x, y) => (x instanceof Date && y instanceof Date) ? x.getTime() === y.getTime() : String(x) === String(y);
  const res = [];
  REF_ABAS.forEach(nome => {
    const src = ss.getSheetByName(nome);
    if (!src) { res.push(nome + ': já estava separada'); return; }
    const velha = dest.getSheetByName(nome); if (velha) dest.deleteSheet(velha);
    const dst = src.copyTo(dest).setName(nome);
    const n = src.getLastRow(), w = src.getLastColumn();
    if (dst.getLastRow() !== n) throw new Error(nome + ': cópia com ' + dst.getLastRow() + ' linhas, original ' + n + '. Nada foi apagado.');
    const a = src.getRange(1, 1, n, w).getValues(), b = dst.getRange(1, 1, n, w).getValues();
    for (let i = 0; i < n; i++) for (let j = 0; j < w; j++) if (!igual(a[i][j], b[i][j])) throw new Error(nome + ': diferença na linha ' + (i + 1) + '. Nada foi apagado.');
    dst.hideSheet();
    res.push(nome + ': ' + n + ' linhas conferidas');
  });
  dest.getSheets().filter(s => REF_ABAS.indexOf(s.getName()) < 0 && s.getName() !== 'Sobre').forEach(s => dest.deleteSheet(s));
  REF_SS_ = dest;
  CacheService.getScriptCache().remove('bamca_idx');
  if (!JSON.parse(getBamca('Lesney', '01A1')).ok) throw new Error('Teste de leitura do BAMCA falhou na planilha nova. Abas originais mantidas.');
  if (!refGet_('fc:carregado')) throw new Error('Teste de leitura do Charlie Mack falhou na planilha nova. Abas originais mantidas.');
  REF_ABAS.forEach(n => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });
  DriveApp.getFileById(dest.getId()).makeCopy('Cópia Referência (Charlie Mack e BAMCA)', folder_('pasta_backup_id', 'Matchbox • Backups'));
  Logger.log(res.join(' | ') + ' | planilha: ' + dest.getUrl());
  return res.join(' | ');
}

/** Avisos (não bloqueiam) ao editar à mão as colunas técnicas e as abas internas. O app continua gravando normalmente. */
function aplicarProtecoes() {
  const ss = ss_(), DESC = 'Matchbox: ';
  const tirar = s => s.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(s.getProtections(SpreadsheetApp.ProtectionType.SHEET))
    .forEach(p => { if (String(p.getDescription()).indexOf(DESC) === 0) p.remove(); });
  tirar(ss);
  Object.keys(SCHEMA).forEach(n => {
    const sh = sheet_(n);
    sh.getRange(1, APP.TECH_COL, sh.getMaxRows(), TECH_FIELDS.length).protect()
      .setDescription(DESC + 'colunas técnicas (ID, fotos, controle do app)').setWarningOnly(true);
  });
  ['_Log', '_Lixeira', '_Config'].forEach(n => {
    const sh = ss.getSheetByName(n); if (!sh) return;
    sh.protect().setDescription(DESC + 'aba interna do app').setWarningOnly(true);
    sh.hideSheet();
  });
  const ref = refSs_();
  if (ref.getId() !== ss.getId()) { tirar(ref); ref.getSheets().forEach(sh => sh.protect().setDescription(DESC + 'referência').setWarningOnly(true)); }
  return 'ok';
}

/** Lista suspensa (sem bloquear) nas colunas de conservação e caixa, com os valores já usados pelo Marco. */
function aplicarValidacoes() {
  const res = [];
  Object.keys(SCHEMA).forEach(n => {
    const s = SCHEMA[n], sh = sheet_(n), last = Math.max(sh.getLastRow(), APP.FIRST);
    [['cons', ['C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'M'], 'Conservação: C0 a C9 ou M (mint).'],
     ['caixa', ['A', 'B', 'C', 'D', 'E', 'F'], 'Tipo de caixa (A a F e variantes já usadas).']].forEach(([k, base, ajuda]) => {
      const col = s.fields.find(f => f[0] === k)[1];
      const usados = sh.getRange(APP.FIRST, col, last - APP.FIRST + 1, 1).getDisplayValues().map(r => String(r[0]).trim()).filter(Boolean);
      const lista = Array.from(new Set(base.concat(usados))).sort().slice(0, 450);
      const regra = SpreadsheetApp.newDataValidation().requireValueInList(lista, false).setAllowInvalid(true).setHelpText(ajuda).build();
      sh.getRange(APP.FIRST, col, sh.getMaxRows() - APP.FIRST + 1, 1).setDataValidation(regra);
      res.push(n + ' ' + k + ': ' + lista.length + ' valores');
    });
  });
  return res.join(' | ');
}

/** Toda segunda às 4h: confere IDs e cópias. Dá ID a linhas digitadas direto na planilha; IDs repetidos só são avisados. */
function verificacaoSemanal() {
  const v = verificarIntegridade(false);
  const semId = Object.keys(v.abas).reduce((t, n) => t + v.abas[n].semId.length, 0);
  const dup = Object.keys(v.abas).reduce((t, n) => t + v.abas[n].duplicados.length, 0);
  let corrigidos = 0;
  if (semId && !dup) corrigidos = verificarIntegridade(true).corrigidos;
  const ult = PropertiesService.getScriptProperties().getProperty('ultima_copia');
  let horas = ult ? (Date.now() - new Date(ult).getTime()) / 36e5 : 999, copiaRefeita = false;
  if (horas > 36) { backupDiario(); copiaRefeita = true; horas = 0; }
  const r = { quando: Utilities.formatDate(new Date(), APP.TZ, 'dd/MM/yyyy HH:mm'), semId: semId, repetidos: dup, corrigidos: corrigidos,
    copiaRefeita: copiaRefeita, ok: !dup && semId === corrigidos };
  cfgSet_('verificacao_semanal', JSON.stringify(r));
  return r;
}
function instalarVerificacaoSemanal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'verificacaoSemanal').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('verificacaoSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).inTimezone(APP.TZ).create();
  return verificacaoSemanal();
}

/** Botão "Gerar planilha": cópia limpa só com os dados do Marco (sem colunas técnicas), na pasta Matchbox • Exportações. */
function gerarPlanilha() {
  const pasta = DriveApp.getFolderById(cfgGet_('pasta_export_id'));
  const nome = 'Coleção Matchbox Marco Geus - ' + Utilities.formatDate(new Date(), APP.TZ, 'dd-MM-yyyy HH\'h\'mm');
  const novo = SpreadsheetApp.create(nome);
  Object.keys(SCHEMA).forEach((n, i) => {
    const src = sheet_(n), w = APP.TECH_COL - 1, last = Math.max(src.getLastRow(), APP.HDR_ROW);
    const dst = i === 0 ? novo.getSheets()[0].setName(n) : novo.insertSheet(n);
    if (dst.getMaxRows() < last) dst.insertRowsAfter(dst.getMaxRows(), last - dst.getMaxRows());
    if (dst.getMaxColumns() < w) dst.insertColumnsAfter(dst.getMaxColumns(), w - dst.getMaxColumns());
    const rg = src.getRange(1, 1, last, w);
    dst.getRange(1, 1, last, w).setValues(rg.getValues()).setNumberFormats(rg.getNumberFormats());
    dst.getRange(APP.HDR_ROW, 1, 1, w).setFontWeight('bold');
    dst.setFrozenRows(APP.HDR_ROW);
  });
  const arq = DriveApp.getFileById(novo.getId());
  arq.moveTo(pasta);
  arq.setDescription('Cópia limpa gerada pelo app (sem colunas técnicas). Não é usada pelo app; pode ser editada ou enviada à vontade.');
  log_('gerar planilha', 'Lesney/Superfast', null, { id: '', nome: nome, mack: '' });
  return JSON.stringify({ nome: nome, url: novo.getUrl(), xlsx: 'https://docs.google.com/spreadsheets/d/' + novo.getId() + '/export?format=xlsx' });
}
function gerarPlanilhaPeloMenu() {
  const r = JSON.parse(gerarPlanilha());
  avisar_('Planilha gerada: "' + r.nome + '"\n\nEla está na pasta Matchbox • Exportações.\n' + r.url);
}

/** Rodar uma vez pelo editor: aplica todas as melhorias de estabilidade da v1.11, na ordem segura. */
function instalarEstabilidade() {
  const log = [];
  const passo = (t, f) => { try { log.push('OK   ' + t + ': ' + f()); } catch (e) { log.push('ERRO ' + t + ': ' + e.message); } };
  passo('Pasta de exportações', () => folder_('pasta_export_id', 'Matchbox • Exportações').getId());
  passo('Validações', aplicarValidacoes);
  passo('Separar referência', separarReferencia);
  passo('Proteções (aviso)', aplicarProtecoes);
  passo('Conferência semanal', () => JSON.stringify(instalarVerificacaoSemanal()));
  Logger.log(log.join('\n'));
  return log.join('\n');
}

/** Dá aos editores da planilha principal (o Marco) acesso à referência e à pasta de exportações. Rodar uma vez pelo editor. */
function compartilharEstabilidade() {
  const emails = DriveApp.getFileById(ss_().getId()).getEditors().map(u => u.getEmail()).filter(Boolean);
  const ref = DriveApp.getFileById(cfgGet_('ref_planilha_id')), pasta = DriveApp.getFolderById(cfgGet_('pasta_export_id'));
  emails.forEach(e => { try { ref.addEditor(e); } catch (x) {} try { pasta.addEditor(e); } catch (x) {} });
  aplicarProtecoes();
  const r = 'Compartilhado com: ' + emails.join(', ') + ' | proteções reaplicadas';
  Logger.log(r);
  return r;
}
