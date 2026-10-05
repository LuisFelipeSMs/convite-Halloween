/* ===== Lista de confirmações do convite de Halloween =====
   Cole este código em Extensões → Apps Script da sua planilha
   e publique como App da Web (veja o passo a passo no chat). */

const ABA = 'Confirmações';
const CRIATURAS = ['Vampiro', 'Bruxa', 'Lobisomem', 'Fantasma', 'Zumbi', 'Múmia'];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(['Data', 'Nome', 'Criatura', 'Decisão']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// A página lê a lista
function doGet() {
  const linhas = aba_().getDataRange().getValues().slice(1);
  const dados = linhas
    .filter(l => l[1])
    .map(l => ({
      nome: String(l[1]),
      classe: String(l[2] || ''),
      decisao: l[3] === 'Vai' ? 'sim' : 'nao',
    }));
  return json_(dados);
}

// A página envia uma confirmação
function doPost(e) {
  const d = JSON.parse(e.postData.contents || '{}');
  let nome = String(d.nome || '').trim().slice(0, 60);
  if (!nome) return json_({ ok: false });
  // Impede que um nome vire fórmula na planilha
  if (/^[=+\-@]/.test(nome)) nome = "'" + nome;
  const classe = CRIATURAS.indexOf(d.classe) >= 0 ? d.classe : '';
  const decisao = d.decisao === 'sim' ? 'Vai' : 'Não vai';

  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    aba_().appendRow([new Date(), nome, classe, decisao]);
  } finally {
    trava.releaseLock();
  }
  return json_({ ok: true });
}
