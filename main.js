//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🤖❗ Menú de Parametros')
      .addItem('🔔- Importar Excel y Convertir', 'importLocal')
      .addItem('📄➔📄- Copiar Datos de VT12', 'copyDataFromVT12File')
      .addItem('❎🗑- Limpiar Todo', 'confirmClearData')
      .addToUi();
}

//----------------------------------------------------------
//----------------------------------------------------------
//---Forma #1 de Convertir a Sheets buscandolo por nombre---
//----------------------------------------------------------
//----------------------------------------------------------
function convertExcel_to_GoogleSheets(){
  var files = DriveApp.getFilesByName("Close_Cards_Data.xlsx");

  while(files.hasNext()){
    var archivo = files.next();
    var nombre = archivo.getName();
    var id = archivo.getId();
    var blob = archivo.getBlob();
  }
  
  var folderId = "1_Xkb_TBY63MI1fMdNUQx8zM4jk8jEwyN"; // ID de la carpeta deseada
  var nvaHCG = {
    title: "[a GSheets] " + nombre,
    parents: [{id: folderId}],
    mimeType: MimeType.GOOGLE_SHEETS
  }
  var hcg = Drive.Files.insert(nvaHCG, blob, {convert:true}); //Hoja de Calculo Google

  var titulo = hcg.title ;
  var enlace = hcg.alternateLink;
  
  var htmlOutput = HtmlService
    .createHtmlOutput('<p>Nombre: '+ titulo + '</p>' +
                      '<p>Abrir desde aquí : <a target="_blank" href="'+enlace+'">ver archivo</a></p>')  
    .setWidth(300)
    .setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Conversión exitosa');
}

//----------------------------------------------------------
//----------------------------------------------------------
//---Forma #2 de Convertir a Sheets de un archivo Local-----
//----------------------------------------------------------
//----------------------------------------------------------

//Funcion para Conectarse al Sheet
function conectionSheets() {
    //Conectar Sheets a AppScript
   const sheetId = '19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA';
   const sheet = SpreadsheetApp.openById(sheetId);
    //Conectar Hojas especificas
   const p_Carga = sheet.getSheetByName('Carga');
   const p_Transportadoras = sheet.getSheetByName('TRANSPORTADORAS');
   const p_Km_TipoTransporte = sheet.getSheetByName('Km-Tipo Transporte');
   
   return {sheet, p_Carga, p_Transportadoras, p_Km_TipoTransporte};
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

function copyDataFromVT12File() {
  var archivoDestino = SpreadsheetApp.openById('19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA');
  var destino = archivoDestino.getActiveSheet();
  var hojaDestino = destino.getRange("Carga")



  // var fechaActual = new Date();
  // var mesAnterior = fetchLastMonth(); // Obtener el mes anterior
  // var currentYear = fechaActual.getFullYear(); // Obtener el año actual
  //var nombreArchivoOrigen = "[GSheets-]VT12 " + mesAnterior + " " + currentYear;
  // var nombreArchivoOrigen = "Qwerty123";
  // Logger.log (nombreArchivoOrigen)

  // Logger.log("Nombre del archivo de origen: " + nombreArchivoOrigen);

  // var archivoOrigen = SpreadsheetApp.getActive().getSheetByName(nombreArchivoOrigen);

  // if (!archivoOrigen) {
  //   Logger.log("Error: No se encontró el archivo de origen.");
  // }

  // var hojaOrigen = archivoOrigen.getActiveSheet();
  // var datosColumnaA = hojaOrigen.getRange("A2:A").getValues().filter(String); // Obtener los datos de la columna A sin filas vacías
  
  // // Abrir el archivo de destino y pegar los datos en la columna B a partir de la fila 18
  // if (!archivoDestino) {
  //   Logger.log("Error: No se encontró el archivo de destino.");
  // }

  // var hojaDestino = archivoDestino.getActiveSheet();
  // var rangoDestino = hojaDestino.getRange(18, 2, datosColumnaA.length, 1); // Definir el rango de destino (columna B, fila 18 en adelante)
  // rangoDestino.setValues(datosColumnaA); // Pegar los datos en el rango de destino
}


function fetchLastMonth() {
  var currentDate = new Date();
  var currentMonth = currentDate.getMonth();

  var previousMonth = (currentMonth === 0) ? 11 : currentMonth - 1; // Restar 1 al mes actual // Si es enero (mes 0), el mes anterior es diciembre (mes 11)
  
  // Obtener el nombre del mes anterior en español y Mayus
  var monthNames = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
  ];
  var previousMonthName = monthNames[previousMonth];
  //Logger.log("Mes anterior: " + previousMonthName);
  return previousMonthName;
}


