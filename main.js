//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🤖❗ Menú de Parametros')
      .addItem('🔔- Importar Excel y Convertir', 'importLocal')
      .addItem('📄➔📄- Copiar Datos de VT12', 'copyDataFromVT12File')
      .addItem('❌✅ - Eliminar Filas por Condiciones', 'removeSpecificRows')
      .addItem('Depurar Base de Datos', '')
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
    confirmAndCleanData('Carga', '¿Está seguro de que desea limpiar los datos de "Carga"?\n\nEste proceso limpiará cualquier tipo de dato', 'U');
}

function copyDataFromVT12File() {
  //Datos del Archivo Origen
  var fechaActual = new Date();
  var mesAnterior = fetchLastMonth(); // Obtener el mes anterior
  var currentYear = fechaActual.getFullYear(); // Obtener el año actual
  var nombreArchivoOrigen = "[GSheets-]VT12 " + mesAnterior + " " + currentYear;
  Logger.log (nombreArchivoOrigen)
  var nombreHojaOrigen = "Hoja1";
  var rangoDatosOrigen = "A2:A";
  var rangoCategoriasOrigen = "M2:M";
  var rangoACOrigen = "AC2:AC";
  var rangoFOrigen = "F2:F";
  var rangoXOrigen = "X2:X";
  var rangoLOrigen = "L2:L";
  var rangoBOrigen = "B2:B";
  var rangoAVOrigen = "AV2:AV";
  var rangoHOrigen = "H2:H"

  //Datos del Archivo Destino
  var idArchivoDestino = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  var nombreHojaDestino = "Carga";
  var filaInicioDestino = 18;
  var columnaDestino = 2;

  var archivosOrigen = DriveApp.getFilesByName(nombreArchivoOrigen);
  if (archivosOrigen.hasNext()) {
    var archivoOrigen = archivosOrigen.next();
    var hojaOrigen = SpreadsheetApp.openById(archivoOrigen.getId()).getSheetByName(nombreHojaOrigen);
    var datosOrigen = hojaOrigen.getRange(rangoDatosOrigen).getValues();
    var categoriasOrigen = hojaOrigen.getRange(rangoCategoriasOrigen).getValues();
    var acOrigen = hojaOrigen.getRange(rangoACOrigen).getValues();
    var fOrigen = hojaOrigen.getRange(rangoFOrigen).getValues();
    var xOrigen = hojaOrigen.getRange(rangoXOrigen).getValues();
    var lOrigen = hojaOrigen.getRange(rangoLOrigen).getValues();
    var bOrigen = hojaOrigen.getRange(rangoBOrigen).getValues();
    var avOrigen = hojaOrigen.getRange(rangoAVOrigen).getValues();
    var hOrigen = hojaOrigen.getRange(rangoHOrigen).getValues();

    // Acceder al archivo de destino
    var archivoDestino = SpreadsheetApp.openById(idArchivoDestino);
    var hojaDestino = archivoDestino.getSheetByName(nombreHojaDestino);

    // Filtrar los datos de la columna B que no comiencen por "580"
    var datosFiltrados = [];
    for (var i = 0; i < bOrigen.length; i++) {
      if (!bOrigen[i][0] || bOrigen[i][0].toString().indexOf("580") !== 0) {
        datosFiltrados.push(datosOrigen[i]);
      }
    }

    // Calcular la cantidad de filas de datos a copiar
    var numRows = datosFiltrados.length;

    // Pegar los datos en la hoja de destino
    hojaDestino.getRange(filaInicioDestino, columnaDestino, numRows, 1).setValues(datosFiltrados);

    // Agregar "Pastas" en la columna A y "Seco" en la columna F
    hojaDestino.getRange(filaInicioDestino, 1, numRows, 1).setValue("Pastas");
    hojaDestino.getRange(filaInicioDestino, 6, numRows, 1).setValue("Seco");

    // Verificar las categorías y colocar "Primario" o "Secundario" en la columna C
    for (var i = 0; i < categoriasOrigen.length; i++) {
      if (categoriasOrigen[i][0] === "ZP01" || categoriasOrigen[i][0] === "ZP02" || categoriasOrigen[i][0] === "ZP07") {
        hojaDestino.getRange(filaInicioDestino + i, 3).setValue("Primario");
      } else if (categoriasOrigen[i][0] === "ZP03" || categoriasOrigen[i][0] === "ZP04" || categoriasOrigen[i][0] === "ZP05" || categoriasOrigen[i][0] === "ZP06" || categoriasOrigen[i][0] === "ZP08") {
        hojaDestino.getRange(filaInicioDestino + i, 3).setValue("Secundario");
      }
    }

    // Copiar datos adicionales del archivo de origen al archivo de destino
    hojaDestino.getRange(filaInicioDestino, 17, numRows, 1).setValues(acOrigen);
    hojaDestino.getRange(filaInicioDestino, 4, numRows, 1).setValues(fOrigen);
    hojaDestino.getRange(filaInicioDestino, 5, numRows, 1).setValues(xOrigen);
    hojaDestino.getRange(filaInicioDestino, 20, numRows, 1).setValues(lOrigen);
    hojaDestino.getRange(filaInicioDestino, 19, numRows, 1).setValues(bOrigen);
    hojaDestino.getRange(filaInicioDestino, 12, numRows, 1).setValues(avOrigen);
    hojaDestino.getRange(filaInicioDestino, 21, numRows, 1).setValues(hOrigen);
  } else {
    Logger.log("¡No se encontró el archivo de origen!");
  }
}

//Funcion para calcular el mes perteneciente al nombre del Excel
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

//Funcion dedicada a eliminar las filas especificas que no se envian (Limpieza de la BD )
function removeSpecificRows (){
 // ID del archivo en el que se trabajarán los datos
 var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  
 // Abrir el archivo de destino
 var archivoDestino = SpreadsheetApp.openById(idArchivo);
 var hojaDestino = archivoDestino.getSheetByName("Carga");

 // Obtener los datos de la hoja
 var datos = hojaDestino.getDataRange().getValues();
 
 // Recorrer los datos desde la última fila hasta la primera
 for (var i = datos.length - 1; i >= 0; i--) {
   var fila = datos[i];
   
   // Eliminar filas que comiencen con "580" en la columna B
   if (fila[1].toString().indexOf("580") === 0) {
     hojaDestino.deleteRow(i + 1);
   }
   
   // Eliminar filas donde el valor en la columna S sea "PINTER"
   if (fila[18].toString() === "PINTER") {
     hojaDestino.deleteRow(i + 1);
   }
   
   // Eliminar filas donde el valor en la columna T sea "DR11" o "DR15"
   if (fila[19].toString() === "DR11" || fila[19].toString() === "DR15") {
     hojaDestino.deleteRow(i + 1);
   }
 }
}


