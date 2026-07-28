function irParaDataAtual() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Nome da aba específica (substitua pelo nome correto)
  const sheetName = "PAGAMENTOS"; 
  const sheet = spreadsheet.getSheetByName(sheetName);

  // Verifica se a aba existe
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`A aba "${sheetName}" não foi encontrada!`);
    return;
  }

  // Configurações (ajuste conforme sua planilha)
  const timeZone = Session.getScriptTimeZone(); // Pega o fuso horário da planilha
  const dateFormat = "dd/MM/yyyy"; // Formato exibido na planilha
  const dateColumn = "A"; // Coluna das datas

  // Formata a data de hoje conforme o exibido na planilha
  const today = Utilities.formatDate(new Date(), timeZone, dateFormat);
  
  // Procura na coluna das datas
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(`${dateColumn}1:${dateColumn}${lastRow}`);
  const dates = range.getDisplayValues(); // Valores como texto

  // Loop para encontrar a data
  for (let i = 0; i < dates.length; i++) {
    if (dates[i][0] === today) {
      sheet.activate(); // Ativa a aba específica
      sheet.getRange(i + 1, 1).activate(); // Ativa a célula
      return;
    }
  }

  SpreadsheetApp.getUi().alert("Data de hoje não encontrada na aba " + sheetName + "!");
}