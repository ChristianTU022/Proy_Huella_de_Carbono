//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🤖❗ Menú de Parametros')
      .addItem('🔔- Importar Excel Local y Convertir', 'importLocal')
      .addItem('❎🗑- Limpiar Todo', 'confirmClearData')
      .addToUi();
}

//Funcion para Conectarse al Sheet
function conectionSheets() {
    //Conectar Sheets a AppScript
   const sheetId = '19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA';
   const sheet = SpreadsheetApp.openById(sheetId);
    //Conectar Hojas especificas
   const p_Carga = sheet.getSheetByName('Carga');
   
   return { sheet, p_Carga};
}

//Funcion Para Confirmar Limpieza de Datos
function confirmAndCleanData(sheetName, confirmationMessage, lastColumn) {
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert(
      'Confirmación',
      confirmationMessage,
      ui.ButtonSet.YES_NO);
  
    if (respuesta == ui.Button.YES) {
      const { sheet } = conectionSheets();
      const targetSheet = sheet.getSheetByName(sheetName);
      const lastRow = targetSheet.getLastRow();
      const range = 'A18:' + lastColumn + lastRow;
      targetSheet.getRange(range).clearContent();
    }
  }
  

//Funcion que permite Limpiar los datos del formulario sheets la hoja "Carga"
function confirmClearData() {
    //Se debe especificar hasta el numero de Columna que se desea eliminar (ultimo parametro)
    confirmAndCleanData('Carga', '¿Está seguro de que desea limpiar los datos de "Carga"?\n\nEste proceso limpiará cualquier tipo de dato', 'R');
}
  





