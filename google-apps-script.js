// ============================================================
// STIMULUS — Google Apps Script
// Cole este código em: Extensions > Apps Script
// Depois: Deploy > New deployment > Web App
//   Execute as: Me
//   Who has access: Anyone
// Copie a URL gerada e cole em index.html e admin.html
// ============================================================

const COLS = [
  '_data', '_id', 'nome', 'instagram', 'whatsapp', 'sexo', 'idade',
  'altura', 'peso', 'fisico_descricao', 'fisico_sentimento', 'objetivo',
  'experiencia', 'dias_treino', 'local_treino', 'lesao', 'qual_lesao',
  'medicamento', 'qual_medicamento', 'sono', 'estresse', 'alimentacao',
  'restricao', 'qual_restricao', 'aparelhos_gosta', 'aparelhos_falta',
  'exercicios_nao_gosta', 'observacoes'
]

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let sheet = ss.getSheetByName('Respostas')
    if (!sheet) sheet = ss.insertSheet('Respostas')

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLS)
      sheet.getRange(1, 1, 1, COLS.length).setFontWeight('bold').setBackground('#0000FF').setFontColor('#FFFFFF')
      sheet.setFrozenRows(1)
    }

    const data = JSON.parse(e.postData.contents)
    sheet.appendRow(COLS.map(col => data[col] !== undefined ? data[col] : ''))

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON)
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON)
  }
}

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respostas')
    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON)
    }

    const data = sheet.getDataRange().getValues()
    const headers = data[0]
    const rows = data.slice(1).map(row => {
      const obj = {}
      headers.forEach((h, i) => { if (h) obj[h] = row[i] })
      return obj
    })

    return ContentService
      .createTextOutput(JSON.stringify(rows))
      .setMimeType(ContentService.MimeType.JSON)
  } catch (err) {
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON)
  }
}
