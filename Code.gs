/**
 * SAFE Escola de Aviação - Controle de Presença
 * Curso: Preparatório Tripulante Azul
 *
 * Backend Google Apps Script (Web App)
 * Todas as operações via GET para evitar CORS.
 *
 * SETUP:
 * 1. Crie uma planilha no Google Sheets
 * 2. Extensões > Apps Script > cole este código
 * 3. Execute a função setup() uma vez (autorize quando pedir)
 * 4. Implantar > Nova implantação > App da Web
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 * 5. Copie a URL gerada e cole na constante API_URL do index.html
 */

var ABA_ALUNOS = 'ALUNOS';
var ABA_AULAS = 'AULAS';
var ABA_PRESENCAS = 'PRESENCAS';
var ABA_INSTRUTORES = 'INSTRUTORES';

/* ============================================================
 * SETUP - executar manualmente uma única vez
 * ============================================================ */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(ABA_ALUNOS)) {
    var sa = ss.insertSheet(ABA_ALUNOS);
    sa.getRange(1, 1, 1, 4).setValues([['ID', 'NOME', 'CANAC', 'STATUS']]);
    sa.getRange(2, 1, 3, 4).setValues([
      ['A001', 'Aluno Exemplo 1', '123456', 'ATIVO'],
      ['A002', 'Aluno Exemplo 2', '', 'ATIVO'],
      ['A003', 'Aluno Exemplo 3', '654321', 'ATIVO']
    ]);
    formatarCabecalho(sa, 4);
  }

  if (!ss.getSheetByName(ABA_AULAS)) {
    var su = ss.insertSheet(ABA_AULAS);
    su.getRange(1, 1, 1, 5).setValues([['ID', 'DATA', 'TEMA', 'INSTRUTOR', 'STATUS']]);
    formatarCabecalho(su, 5);
  }

  if (!ss.getSheetByName(ABA_PRESENCAS)) {
    var sp = ss.insertSheet(ABA_PRESENCAS);
    sp.getRange(1, 1, 1, 4).setValues([['ID_AULA', 'ID_ALUNO', 'STATUS', 'TIMESTAMP']]);
    formatarCabecalho(sp, 4);
  }

  if (!ss.getSheetByName(ABA_INSTRUTORES)) {
    var si = ss.insertSheet(ABA_INSTRUTORES);
    si.getRange(1, 1, 1, 2).setValues([['NOME', 'STATUS']]);
    si.getRange(2, 1, 2, 2).setValues([
      ['Instrutor Exemplo 1', 'ATIVO'],
      ['Instrutor Exemplo 2', 'ATIVO']
    ]);
    formatarCabecalho(si, 2);
  }

  // Remove a aba padrão se ainda existir
  var padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 4) ss.deleteSheet(padrao);
}

function formatarCabecalho(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setBackground('#1D2951')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/* ============================================================
 * ROTEADOR
 * ============================================================ */
function doGet(e) {
  var action = (e.parameter.action || '').toString();
  var result;

  try {
    switch (action) {
      case 'init':
        result = getInit();
        break;
      case 'criarAula':
        result = criarAula(e.parameter);
        break;
      case 'salvarChamada':
        result = salvarChamada(e.parameter);
        break;
      case 'excluirAula':
        result = excluirAula(e.parameter);
        break;
      default:
        result = { ok: false, erro: 'Ação inválida: ' + action };
    }
  } catch (err) {
    result = { ok: false, erro: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 * AÇÕES
 * ============================================================ */

// Retorna tudo em uma única chamada: alunos, aulas e presenças
function getInit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var alunos = lerAba(ss, ABA_ALUNOS).map(function (r) {
    return { id: String(r[0]), nome: String(r[1]), canac: String(r[2] || ''), status: String(r[3] || 'ATIVO') };
  }).filter(function (a) { return a.id && a.status.toUpperCase() === 'ATIVO'; });

  var aulas = lerAba(ss, ABA_AULAS).map(function (r) {
    return {
      id: String(r[0]),
      data: formatarData(r[1]),
      tema: String(r[2] || ''),
      instrutor: String(r[3] || ''),
      status: String(r[4] || 'AGENDADA')
    };
  }).filter(function (a) { return a.id; });

  var presencas = lerAba(ss, ABA_PRESENCAS).map(function (r) {
    return { idAula: String(r[0]), idAluno: String(r[1]), status: String(r[2]) };
  }).filter(function (p) { return p.idAula; });

  var instrutores = lerAba(ss, ABA_INSTRUTORES).map(function (r) {
    return { nome: String(r[0] || '').trim(), status: String(r[1] || 'ATIVO') };
  }).filter(function (i) { return i.nome && i.status.toUpperCase() === 'ATIVO'; })
    .map(function (i) { return i.nome; });

  return { ok: true, alunos: alunos, aulas: aulas, presencas: presencas, instrutores: instrutores };
}

// Cria nova aula. Params: data (YYYY-MM-DD), tema, instrutor
function criarAula(p) {
  if (!p.data || !p.tema) return { ok: false, erro: 'Data e tema são obrigatórios.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ABA_AULAS);
  var id = 'AULA-' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyyMMddHHmmss');

  sheet.appendRow([id, p.data, p.tema, p.instrutor || '', 'AGENDADA']);
  return { ok: true, id: id };
}

// Salva/atualiza a chamada de uma aula.
// Params: idAula, dados (JSON: [{idAluno, status}, ...])
function salvarChamada(p) {
  if (!p.idAula || !p.dados) return { ok: false, erro: 'Parâmetros incompletos.' };

  var dados = JSON.parse(p.dados);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ABA_PRESENCAS);
  var ts = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

  // Remove registros anteriores da aula (permite refazer a chamada)
  var valores = sheet.getDataRange().getValues();
  for (var i = valores.length - 1; i >= 1; i--) {
    if (String(valores[i][0]) === String(p.idAula)) {
      sheet.deleteRow(i + 1);
    }
  }

  // Grava os novos registros
  if (dados.length > 0) {
    var linhas = dados.map(function (d) {
      return [p.idAula, d.idAluno, d.status, ts];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, linhas.length, 4).setValues(linhas);
  }

  // Marca a aula como realizada
  var sheetAulas = ss.getSheetByName(ABA_AULAS);
  var aulas = sheetAulas.getDataRange().getValues();
  for (var j = 1; j < aulas.length; j++) {
    if (String(aulas[j][0]) === String(p.idAula)) {
      sheetAulas.getRange(j + 1, 5).setValue('REALIZADA');
      break;
    }
  }

  return { ok: true };
}

// Exclui uma aula e suas presenças. Params: idAula
function excluirAula(p) {
  if (!p.idAula) return { ok: false, erro: 'ID da aula não informado.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetAulas = ss.getSheetByName(ABA_AULAS);
  var aulas = sheetAulas.getDataRange().getValues();
  for (var i = aulas.length - 1; i >= 1; i--) {
    if (String(aulas[i][0]) === String(p.idAula)) sheetAulas.deleteRow(i + 1);
  }

  var sheetPres = ss.getSheetByName(ABA_PRESENCAS);
  var pres = sheetPres.getDataRange().getValues();
  for (var j = pres.length - 1; j >= 1; j--) {
    if (String(pres[j][0]) === String(p.idAula)) sheetPres.deleteRow(j + 1);
  }

  return { ok: true };
}

/* ============================================================
 * AUXILIARES
 * ============================================================ */
function lerAba(ss, nome) {
  var sheet = ss.getSheetByName(nome);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function formatarData(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd');
  }
  return String(v);
}
