/**
 * Controle de Ajoduns – Ilé Òpó Iná Àse Oya
 * Servidor (Google Apps Script). Os dados ficam nas abas da própria planilha.
 * Atende o site antigo (link do Apps Script) e o aplicativo (GitHub Pages) ao mesmo tempo.
 */

// >>> TROQUE a senha abaixo por uma que só os diretores saibam <<<
var PASSWORD = 'oya2027';

var SHEETS = {
  Celebracoes:   ['id', 'name', 'date', 'note'],
  Gastos:        ['id', 'ajodunId', 'desc', 'cat', 'amount', 'createdAt'],
  Contribuicoes: ['id', 'ajodunId', 'person', 'amount', 'paid', 'createdAt'],
  Filhos:        ['id', 'name', 'valor', 'createdAt', 'inicio', 'nascimento', 'feitura', 'foto', 'cobranca', 'telefone', 'tratamento'],
  Mensalidades:  ['id', 'filhoId', 'year', 'month', 'amount', 'paid', 'paidAt', 'justificou', 'obs'],
  Config:        ['id', 'value'],
  Solicitacoes:  ['id', 'name', 'createdAt', 'nascimento', 'feitura', 'filhoId'],
  // Comprovantes de pagamento. tipo = 'mens' (refId = id da mensalidade, ex.: iuri_2026-11)
  // ou 'contrib' (refId = id da contribuição). Só um comprovante por refId.
  Anexos:        ['id', 'refId', 'tipo', 'fileId', 'nome', 'mime', 'createdAt'],
  // Checklist da iniciação. lista = 'lit' (material litúrgico) ou 'merc' (materiais de mercado).
  Checklist:     ['id', 'lista', 'item', 'qtd', 'obs', 'createdAt']
};

var SEED_CELEBRACOES = [
  ['oxala',  'Ajodún de Oxalá', '2027-03-13', ''],
  ['ogun',   'Ajodún de Ògún', '2027-05-22', ''],
  ['oxossi', 'Ajodún de Oxossi + Odun Eje Iuri de Oxossi', '2027-06-26', 'Duas cerimônias no mesmo dia.'],
  ['oya',    'Ajodún de Oya – Iyalorisá', '2027-10-23', ''],
  ['ibeji',  'Ajodún de Ibeji', '2027-10-24', ''],
  ['rosa',   'Rosa Vermelha e Sete Encruzilhadas', '2027-11-27', '']
];

var SEED_FILHOS = [
  ['iuri',     'Iuri de Oxossi', 0, 1],
  ['katia',    'Katia de Iyemanjá', 0, 2],
  ['daniele',  'Daniele de Iyemanjá', 0, 3],
  ['daiana',   'Daiana de Xangô', 0, 4],
  ['fernando', 'Fernando de Oxalá', 0, 5],
  ['rafael',   'Rafael de Oxalá', 0, 6]
];

/**
 * Sem parâmetros: abre o site antigo (Index). Com ?cadastro=1: formulário público de cadastro.
 * O aplicativo (GitHub Pages) conversa por doPost, usando a mesma senha dos diretores.
 */
function doGet(e) {
  var cadastro = e && e.parameter && e.parameter.cadastro;
  return HtmlService.createHtmlOutputFromFile(cadastro ? 'Cadastro' : 'Index')
    .setTitle(cadastro ? 'Cadastro de filhos' : 'Controle de Ajoduns')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Porta de entrada do aplicativo: recebe { pass, action, p } e devolve o mesmo JSON da função api(). */
function doPost(e) {
  var corpo = {};
  try { corpo = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { corpo = {}; }
  var resposta = api(corpo.pass, corpo.action, corpo.p);
  if (resposta.indexOf('"error":"auth"') > -1) Utilities.sleep(1500); // atrasa quem tenta adivinhar a senha
  return ContentService.createTextOutput(resposta).setMimeType(ContentService.MimeType.JSON);
}

/** Única porta de entrada chamada pela tela. Devolve sempre um texto JSON. */
function api(pass, action, p) {
  try {
    if (pass !== PASSWORD) return JSON.stringify({ ok: false, error: 'auth' });
    p = p || {};
    if (action === 'getAll') {
      ensure_();
      return JSON.stringify({ ok: true, data: readAll_() });
    }
    if (action === 'getFotos') {
      ensure_();
      return JSON.stringify({ ok: true, data: getFotos_() });
    }
    if (action === 'getComp') {
      return JSON.stringify({ ok: true, data: getComp_(p.id) });
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (action === 'put') put_(p.sheet, p.row);
      else if (action === 'del') del_(p.sheet, p.id);
      else if (action === 'delWhere') delWhere_(p.sheet, p.field, p.value);
      else if (action === 'setFoto') setFoto_(p.id, p.foto);
      else if (action === 'addComp') addComp_(p);
      else throw new Error('Ação desconhecida');
    } finally {
      lock.releaseLock();
    }
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ ok: false, error: String(e && e.message ? e.message : e) });
  }
}

function ensure_() {
  var ss = SpreadsheetApp.getActive();
  var created = false;
  Object.keys(SHEETS).forEach(function (name) {
    if (ss.getSheetByName(name)) { fixHeaders_(ss.getSheetByName(name), name); return; }
    var h = SHEETS[name];
    var sh = ss.insertSheet(name);
    created = true;
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    var seed = null;
    if (name === 'Celebracoes') { sh.getRange('C:C').setNumberFormat('@'); seed = SEED_CELEBRACOES; }
    if (name === 'Filhos') { sh.getRange('E:G').setNumberFormat('@'); sh.getRange('J:J').setNumberFormat('@'); seed = SEED_FILHOS; }
    if (name === 'Solicitacoes') sh.getRange('D:F').setNumberFormat('@');
    if (name === 'Anexos') sh.getRange('B:D').setNumberFormat('@');
    if (name === 'Checklist') sh.getRange('B:E').setNumberFormat('@');
    if (seed) sh.getRange(2, 1, seed.length, seed[0].length).setValues(seed);
  });
  if (created) {
    ['Página1', 'Sheet1', 'Planilha1'].forEach(function (n) {
      var s = ss.getSheetByName(n);
      if (s && ss.getSheets().length > 1 && s.getLastRow() <= 1 && s.getLastColumn() <= 1) ss.deleteSheet(s);
    });
  }
}

/** Planilhas criadas numa versão antiga ganham as colunas novas sem perder dados. */
function fixHeaders_(sh, name) {
  var h = SHEETS[name];
  var cur = sh.getRange(1, 1, 1, h.length).getValues()[0];
  var diff = false;
  for (var i = 0; i < h.length; i++) if (cur[i] !== h[i]) diff = true;
  if (!diff) return;
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  if (name === 'Filhos') { sh.getRange('E:G').setNumberFormat('@'); sh.getRange('J:J').setNumberFormat('@'); }
  if (name === 'Solicitacoes') sh.getRange('D:F').setNumberFormat('@');
  if (name === 'Anexos') sh.getRange('B:D').setNumberFormat('@');
  if (name === 'Checklist') sh.getRange('B:E').setNumberFormat('@');
}

function readAll_() {
  var ss = SpreadsheetApp.getActive();
  var tz = Session.getScriptTimeZone();
  var out = {};
  Object.keys(SHEETS).forEach(function (name) {
    var h = SHEETS[name];
    var sh = ss.getSheetByName(name);
    var rows = [];
    var last = sh.getLastRow();
    if (last > 1) {
      var values = sh.getRange(2, 1, last - 1, h.length).getValues();
      values.forEach(function (r) {
        if (r[0] === '' || r[0] === null) return;
        var o = {};
        h.forEach(function (k, i) {
          var v = r[i];
          if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
          o[k] = v;
        });
        rows.push(o);
      });
    }
    out[name] = rows;
  });
  return out;
}

function sheet_(name) {
  if (!SHEETS.hasOwnProperty(name)) throw new Error('Aba inválida');
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) { ensure_(); sh = SpreadsheetApp.getActive().getSheetByName(name); }
  return sh;
}

function clean_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string') {
    v = v.slice(0, 500);
    if (/^[=+\-@]/.test(v)) v = "'" + v;
  }
  return v;
}

function validId_(id) {
  if (typeof id !== 'string' || !/^[A-Za-z][A-Za-z0-9_\-:.]{0,79}$/.test(id)) throw new Error('Identificador inválido');
  return id;
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

/** Grava ou atualiza uma linha. Na aba Filhos, se a foto não vier no pedido, mantém a que já existe. */
function put_(name, row) {
  var sh = sheet_(name);
  var h = SHEETS[name];
  validId_(row && row.id);
  var r = findRow_(sh, row.id);
  var atual = r === -1 ? null : sh.getRange(r, 1, 1, h.length).getValues()[0];
  var arr = h.map(function (k, i) {
    if (name === 'Filhos' && (k === 'foto' || k === 'cobranca' || k === 'telefone' || k === 'tratamento') && row[k] === undefined) return atual ? atual[i] : '';
    return clean_(row[k]);
  });
  if (r === -1) r = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(r, 1, 1, h.length).setValues([arr]);
}

/** Apaga uma linha. Na aba Filhos apaga também a foto no Drive; na aba Anexos, o comprovante. */
function del_(name, id) {
  var sh = sheet_(name);
  validId_(id);
  var r = findRow_(sh, id);
  if (r > 1) {
    if (name === 'Filhos') {
      var f = String(sh.getRange(r, SHEETS.Filhos.indexOf('foto') + 1).getValue() || '');
      if (f) trashFile_(f);
    }
    if (name === 'Anexos') {
      var cf = String(sh.getRange(r, SHEETS.Anexos.indexOf('fileId') + 1).getValue() || '');
      if (cf) trashFile_(cf);
    }
    sh.deleteRow(r);
  }
}

function delWhere_(name, field, value) {
  var sh = sheet_(name);
  var h = SHEETS[name];
  var col = h.indexOf(field);
  if (col < 1) throw new Error('Campo inválido');
  var last = sh.getLastRow();
  if (last < 2) return;
  var values = sh.getRange(2, 1, last - 1, h.length).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][col]) === String(value)) sh.deleteRow(i + 2);
  }
}

/* ------------------------------------------------------------------
 * Fotos dos filhos (ficam numa pasta privada do Drive)
 * ------------------------------------------------------------------ */

var FOTO_FOLDER_KEY = 'FOTO_FOLDER_ID';

function fotoFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(FOTO_FOLDER_KEY);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Controle de Ajoduns - Fotos');
  props.setProperty(FOTO_FOLDER_KEY, f.getId());
  return f;
}

/**
 * Rode UMA VEZ pelo editor (botão Executar) para o Google pedir a permissão do Drive.
 * Depois publique uma nova versão da implantação.
 */
function autorizarDrive() {
  var f = fotoFolder_();
  Logger.log('Pasta de fotos pronta: ' + f.getUrl());
}

function trashFile_(id) {
  try { DriveApp.getFileById(String(id)).setTrashed(true); } catch (e) {}
}

function setFoto_(filhoId, dataUrl) {
  validId_(filhoId);
  var prefix = 'data:image/jpeg;base64,';
  dataUrl = String(dataUrl || '');
  if (dataUrl.indexOf(prefix) !== 0 || dataUrl.length > 400000) throw new Error('Foto inválida ou grande demais');
  var sh = sheet_('Filhos');
  var r = findRow_(sh, filhoId);
  if (r < 2) throw new Error('Filho não encontrado');
  var col = SHEETS.Filhos.indexOf('foto') + 1;
  var old = String(sh.getRange(r, col).getValue() || '');
  var blob = Utilities.newBlob(Utilities.base64Decode(dataUrl.slice(prefix.length)), 'image/jpeg', filhoId + '.jpg');
  var file = fotoFolder_().createFile(blob);
  sh.getRange(r, col).setValue(file.getId());
  if (old) trashFile_(old);
}

/** Devolve {idDoFilho: 'data:image/jpeg;base64,...'}. As fotos ficam privadas no Drive. */
function getFotos_() {
  var sh = sheet_('Filhos');
  var h = SHEETS.Filhos;
  var last = sh.getLastRow();
  var out = {};
  if (last < 2) return out;
  var ci = h.indexOf('foto');
  sh.getRange(2, 1, last - 1, h.length).getValues().forEach(function (r) {
    if (!r[ci]) return;
    try {
      var bytes = DriveApp.getFileById(String(r[ci])).getBlob().getBytes();
      out[r[0]] = 'data:image/jpeg;base64,' + Utilities.base64Encode(bytes);
    } catch (e) {}
  });
  return out;
}

/* ------------------------------------------------------------------
 * Comprovantes de pagamento (ficam numa pasta privada do Drive)
 * Um comprovante por mensalidade (filho + mês) e um por contribuição.
 * Anexar de novo no mesmo item substitui o anterior.
 * ------------------------------------------------------------------ */

var COMP_FOLDER_KEY = 'COMP_FOLDER_ID';

function compFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(COMP_FOLDER_KEY);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Controle de Ajoduns - Comprovantes');
  props.setProperty(COMP_FOLDER_KEY, f.getId());
  return f;
}

function addComp_(p) {
  validId_(p.id);
  var tipo = String(p.tipo || '');
  var ref = validId_(String(p.refId || ''));
  if (tipo === 'mens') {
    var mm = /^(.+)_(\d{4})-(\d{2})$/.exec(ref);
    if (!mm) throw new Error('Mês inválido');
    var y = Number(mm[2]), mo = Number(mm[3]);
    if (y < 2026 || y > 2100 || mo < 1 || mo > 12) throw new Error('Mês inválido');
    if (findRow_(sheet_('Filhos'), mm[1]) < 2) throw new Error('Filho não encontrado');
  } else if (tipo === 'contrib') {
    if (findRow_(sheet_('Contribuicoes'), ref) < 2) throw new Error('Contribuição não encontrada');
  } else {
    throw new Error('Tipo inválido');
  }
  var data = String(p.data || '');
  var m = /^data:(image\/jpeg|application\/pdf);base64,/.exec(data);
  if (!m || data.length > 5000000) throw new Error('Arquivo inválido ou grande demais');
  var mime = m[1];
  var ext = mime === 'application/pdf' ? 'pdf' : 'jpg';
  var blob = Utilities.newBlob(Utilities.base64Decode(data.slice(m[0].length)), mime, 'comprovante-' + ref + '-' + p.id + '.' + ext);
  var file = compFolder_().createFile(blob);

  var sh = sheet_('Anexos');
  var h = SHEETS.Anexos;
  var last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, h.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][1]) === ref && String(vals[i][2]) === tipo) {
        trashFile_(vals[i][3]);
        sh.deleteRow(i + 2);
      }
    }
  }
  var row = [p.id, ref, tipo, file.getId(), clean_(String(p.nome || 'comprovante').slice(0, 80)), mime, Date.now()];
  sh.getRange(Math.max(sh.getLastRow(), 1) + 1, 1, 1, row.length).setValues([row]);
}

/** Devolve {mime, data:'data:...;base64,...'} de um comprovante. */
function getComp_(id) {
  validId_(id);
  var sh = sheet_('Anexos');
  var r = findRow_(sh, id);
  if (r < 2) throw new Error('Comprovante não encontrado');
  var h = SHEETS.Anexos;
  var row = sh.getRange(r, 1, 1, h.length).getValues()[0];
  var mime = String(row[h.indexOf('mime')]);
  var bytes = DriveApp.getFileById(String(row[h.indexOf('fileId')])).getBlob().getBytes();
  return { mime: mime, data: 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes) };
}

/* ------------------------------------------------------------------
 * Cadastro público
 * ------------------------------------------------------------------ */

function normName_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
}

function validDate_(s, minYear) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var p = s.split('-').map(Number);
  var d = new Date(p[0], p[1] - 1, p[2]);
  if (d.getFullYear() !== p[0] || d.getMonth() !== p[1] - 1 || d.getDate() !== p[2]) return false;
  if (p[0] < minYear) return false;
  var lim = new Date();
  lim.setDate(lim.getDate() + 1);
  return d <= lim;
}

/**
 * Chamada pela página pública de cadastro (sem senha).
 * Só grava um pedido na aba Solicitacoes; quem aprova é o administrador.
 * Se o nome já é de um filho cadastrado, o pedido vira uma atualização dos dados dele.
 */
function solicitarCadastro(name, nasc, feitura, trap) {
  try {
    if (trap) return JSON.stringify({ ok: true }); // campo escondido: robô
    name = String(name || '').replace(/\s+/g, ' ').trim();
    nasc = String(nasc || '').trim();
    feitura = String(feitura || '').trim();
    if (name.length < 3 || name.length > 80) return JSON.stringify({ ok: false, error: 'tamanho' });
    if (!validDate_(nasc, 1900)) return JSON.stringify({ ok: false, error: 'nascimento' });
    if (feitura && !validDate_(feitura, 1950)) return JSON.stringify({ ok: false, error: 'feitura' });
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var tipo = 'novo';
    try {
      ensure_();
      var key = normName_(name);
      var ss = SpreadsheetApp.getActive();
      var filhoId = '';
      var fsh = ss.getSheetByName('Filhos');
      var fl = fsh.getLastRow();
      if (fl > 1) {
        fsh.getRange(2, 1, fl - 1, 2).getValues().forEach(function (r) {
          if (normName_(r[1]) === key) filhoId = String(r[0]);
        });
      }
      var sol = ss.getSheetByName('Solicitacoes');
      var sl = sol.getLastRow();
      if (sl > 1) {
        var pend = false;
        sol.getRange(2, 2, sl - 1, 1).getValues().forEach(function (r) {
          if (normName_(r[0]) === key) pend = true;
        });
        if (pend) return JSON.stringify({ ok: false, error: 'pendente' });
      }
      if (sl > 300) return JSON.stringify({ ok: false, error: 'cheio' });
      var id = 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
      sol.getRange(sl + 1, 1, 1, 6).setValues([[id, clean_(name), Date.now(), nasc, feitura, filhoId]]);
      if (filhoId) tipo = 'atualizacao';
    } finally {
      lock.releaseLock();
    }
    return JSON.stringify({ ok: true, tipo: tipo });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'erro' });
  }
}