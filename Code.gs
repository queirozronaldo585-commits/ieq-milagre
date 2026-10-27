// ESCALA DE DEMANDAS - API (Google Apps Script + Planilha Google)
// Use este arquivo como backend da página index.html hospedada na Vercel.
var ABA = "Demandas";
var COLS = ["id", "titulo", "data", "hora", "resp", "status", "criadoPor"];
var EDITAVEIS = ["titulo", "data", "hora", "resp", "status"];

function resposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Lista as demandas
function doGet() {
  try {
    return resposta_({ ok: true, dados: listar() });
  } catch (err) {
    return resposta_({ ok: false, erro: String(err) });
  }
}

// Adiciona, atualiza ou exclui
function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    var r;
    if (p.acao === "adicionar") r = adicionar(p.dados || {});
    else if (p.acao === "atualizar") r = atualizar(p.id, p.campos || {});
    else if (p.acao === "excluir") r = excluir(p.id);
    else return resposta_({ ok: false, erro: "acao invalida" });
    return resposta_({ ok: true, resultado: r });
  } catch (err) {
    return resposta_({ ok: false, erro: String(err) });
  }
}

function aba_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
    sh.getRange("A:G").setNumberFormat("@"); // tudo como texto (evita converter datas)
  }
  return sh;
}

function limpa_(x) {
  return String(x == null ? "" : x).slice(0, 200);
}

function comTrava_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function listar() {
  var sh = aba_();
  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    var o = {};
    for (var j = 0; j < COLS.length; j++) o[COLS[j]] = limpa_(v[i][j]);
    out.push(o);
  }
  return out;
}

function adicionar(d) {
  return comTrava_(function () {
    var sh = aba_();
    var id = Utilities.getUuid().slice(0, 8);
    sh.appendRow([id, limpa_(d.titulo), limpa_(d.data), limpa_(d.hora), "", "Pendente", limpa_(d.criadoPor)]);
    return id;
  });
}

function atualizar(id, campos) {
  return comTrava_(function () {
    var sh = aba_();
    var n = sh.getLastRow() - 1;
    if (n < 1) return false;
    var ids = sh.getRange(2, 1, n, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        var linha = i + 2;
        for (var k in campos) {
          if (EDITAVEIS.indexOf(k) >= 0) sh.getRange(linha, COLS.indexOf(k) + 1).setValue(limpa_(campos[k]));
        }
        return true;
      }
    }
    return false;
  });
}

function excluir(id) {
  return comTrava_(function () {
    var sh = aba_();
    var n = sh.getLastRow() - 1;
    if (n < 1) return false;
    var ids = sh.getRange(2, 1, n, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sh.deleteRow(i + 2);
        return true;
      }
    }
    return false;
  });
}
