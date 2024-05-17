//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🤖❗ Menú de Parametros')
      .addItem('🔔- 1.) Importar Excel y Convertir', 'importLocal')
      .addItem('📄➔📄- 2.) Copiar Datos de VT12', 'copyDataFromVT12File')
      .addItem('❌✅ - 3.) Eliminar Filas por Condiciones', 'removeSpecificRows')
      .addItem('➕ - 4.) Completar Campos Faltantes', 'completeTableFields')
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
    confirmAndCleanData('Carga', '¿Está seguro de que desea limpiar los datos de "Carga"?\n\nEste proceso limpiará cualquier tipo de dato', 'V');
}

function copyDataFromVT12File() {
  //Datos del Archivo Origen
  var fechaActual = new Date();
  var mesAnterior = fetchLastMonth(); // Obtener el mes anterior
  var currentYear = fechaActual.getFullYear(); // Obtener el año actual
  var nombreArchivoOrigen = "[GSheets-]VT12 " + mesAnterior + " " + currentYear;
  //Logger.log(nombreArchivoOrigen);
  var nombreHojaOrigen = "Hoja1";
  var rangoDatosOrigen = "A2:A";
  var rangoCategoriasOrigen = "M2:M";
  var rangoACOrigen = "AC2:AC";
  var rangoFOrigen = "F2:F";
  var rangoXOrigen = "X2:X";
  var rangoLOrigen = "L2:L";
  var rangoBOrigen = "B2:B";
  var rangoAVOrigen = "AV2:AV";
  var rangoHOrigen = "H2:H";
  var rangoOOrigen = "O2:O";
  var rangoPOrigen = "P2:P";
  var rangoQOrigen = "Q2:Q";
  var rangoROrigen = "R2:R";
  var rangoSOrigen = "S2:S";
  var rangoTOrigen = "T2:T";
  var rangoUOrigen = "U2:U";
  var rangoVOrigen = "V2:V";

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
    var oOrigen = hojaOrigen.getRange(rangoOOrigen).getValues();
    var pOrigen = hojaOrigen.getRange(rangoPOrigen).getValues();
    var qOrigen = hojaOrigen.getRange(rangoQOrigen).getValues();
    var rOrigen = hojaOrigen.getRange(rangoROrigen).getValues();
    var sOrigen = hojaOrigen.getRange(rangoSOrigen).getValues();
    var tOrigen = hojaOrigen.getRange(rangoTOrigen).getValues();
    var uOrigen = hojaOrigen.getRange(rangoUOrigen).getValues();
    var vOrigen = hojaOrigen.getRange(rangoVOrigen).getValues();

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
      if (
        categoriasOrigen[i][0] === "ZP01" ||
        categoriasOrigen[i][0] === "ZP02" ||
        categoriasOrigen[i][0] === "ZP07"
      ) {
        hojaDestino.getRange(filaInicioDestino + i, 3).setValue("Primario");
      } else if (
        categoriasOrigen[i][0] === "ZP03" ||
        categoriasOrigen[i][0] === "ZP04" ||
        categoriasOrigen[i][0] === "ZP05" ||
        categoriasOrigen[i][0] === "ZP06" ||
        categoriasOrigen[i][0] === "ZP08"
      ) {
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

    // Verificar si hay datos en las columnas O, P o Q y colocar 1 o 0 en la columna V del destino
    var datosVerificados = [];
    for (var i = 0; i < oOrigen.length; i++) {
      var valorV = oOrigen[i][0] || pOrigen[i][0] || qOrigen[i][0] || 
                  rOrigen[i][0] || sOrigen[i][0] || tOrigen[i][0] || 
                  uOrigen[i][0] || vOrigen[i][0] ? 1 : 0;
      datosVerificados.push([valorV]);
    }
    hojaDestino.getRange(filaInicioDestino, 22, numRows, 1).setValues(datosVerificados);
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

   // Eliminar filas donde el valor en la columna V sea 0
   if (fila[21] === 0) {
    hojaDestino.deleteRow(i + 1);
  }
 }
}


function completeTableFields () {
  //Llamar a la función typeOfBelongingAndFuel "Tipo de Flota o Pertenencia (Dice si es Propio o Contratado)"
  typeOfBelongingAndFuel();
  // Llamar a la función findTypeTransportation para determinar el tipo de transporte que se uso. (SC Sencillo,TB Turbo, TM2 Tractomula 2 ejes, DT Dobletroque, etc)
  findTypeTransportation();
  //Llamar a la funcion amountOfFuelPerTrip para determinar la cantidad de Combustible por Viaje (gal/viaje)
  amountOfFuelPerTrip();
  //Llama a la funcion originDestinRoute, la cual encuentra la ciudad y origen de destino
  originDestinRoute();
  //Llama a la funcion ,la cual calcula el Rendimiento por Viaje (Km/Gal)
  calculateFuelEfficiency();
  //Llamar a la funcion searchNITTransporter, la cual busca el NIT de la Transportadora
  searchNITTransporter();
  //Llamar a la funcion clearColumnsSTUV, la cual despues de llenar todo borra los datos temporales de la columna S, T, U y V
  clearColumnsSTUV();
}

//Funcion para saber si el transporte es propio o contratado y el tipo de Combustible
function typeOfBelongingAndFuel () {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  
  // Abrir el archivo de destino
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaDestino = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de la hoja
  var datos = hojaDestino.getRange("Q18:Q").getValues();
  var numRows = datos.length;
  
  // Recorrer los datos
  for (var i = 0; i < numRows; i++) {
    var valorQ = datos[i][0];
    var valorR = valorQ === "JPO336" ? "Propio" : valorQ ? "Tercerizado" : "";
    
    // Establecer el tipo de combustible según el valor de la pertenencia
    var tipoCombustible = "";
    if (valorR === "Propio") {
      tipoCombustible = "Gas Natural";
    } else if (valorR === "Tercerizado") {
      tipoCombustible = "Diésel";
    }

    // Colocar el tipo de pertenencia en la columna N
    hojaDestino.getRange(18 + i, 14).setValue(tipoCombustible);
    
    hojaDestino.getRange(18 + i, 18).setValue(valorR);
  }
}

//Funcion que con los pesos Dice que tipo de Vehiculo es "Sencillo, Turbo, TM 2 Ejes, TM 3 Ejes, etc"
function findTypeTransportation() {
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaDestino = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de la column L
  var datosL = hojaDestino.getRange("L18:L").getValues();
  var numRows = datosL.length;
  
  for (var i = 0; i < numRows; i++) {
    var valorL = datosL[i][0];

    // Determinar el tipo de vehículo según el valor en la columna L
    var tipoVehiculo = "";
    if (valorL >= 1 && valorL <= 4500) {
      tipoVehiculo = "TB Turbo";
    } else if (valorL >= 4501 && valorL <= 9000) {
      tipoVehiculo = "SC Sencillo";
    } else if (valorL >= 9001 && valorL <= 18000) {
      tipoVehiculo = "DT Dobletroque";
    } else if (valorL >= 18001 && valorL <= 30000) {
      tipoVehiculo = "TM2 Tractomula 2 ejes";
    } else if (valorL > 30000) {
      tipoVehiculo = "TM3 Tractomula 3 ejes";
    }

    // Obtener el texto de la columna U y convertirlo a mayúsculas
    var textoU = hojaDestino.getRange(18 + i, 21).getValue().toUpperCase();

    // Verificar si el texto contiene "MINI"
    if (textoU.indexOf("MINI") !== -1) {
      tipoVehiculo = "MM Minimula";
    }
     
    // Colocar el tipo de vehículo en la columna M
    hojaDestino.getRange(18 + i, 13).setValue(tipoVehiculo);
  }
}

// Esta funcion busca el NIT de la transportadora por su nombre
function searchNITTransporter() {
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener las hojas
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");
  var hojaTransportadoras = archivoDestino.getSheetByName("TRANSPORTADORAS");

  // Obtener los datos de las columnas D en la hoja de "Carga"
  var datosCarga = hojaCarga.getRange("D18:D").getValues();
  var numRows = datosCarga.length;

  // Recorrer los datos
  for (var i = 0; i < numRows; i++) {
    var valorD = datosCarga[i][0];
    
    // Buscar el valor en la hoja de "TRANSPORTADORAS"
    var datosTransportadoras = hojaTransportadoras.getRange("B:B").getValues();
    var index = datosTransportadoras.findIndex(function(row) {
      return row[0] === valorD;
    });

    // Si se encuentra el valor, copiar el valor de la columna C y pegarlo en la hoja de "Carga"
    if (index !== -1) {
      var valorC = hojaTransportadoras.getRange(index + 1, 3).getValue(); // +1 para ajustar al índice base 1
      hojaCarga.getRange(18 + i, 4).setValue(valorC); // Colocar en la columna D de la hoja de "Carga"
    }
  }
}

// Esta funcion tiene el fin de decir cual es la cantidad de combustible por viaje (gal/viaje), con ayuda del Tipo de Vehiculo
function amountOfFuelPerTrip() {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener la hoja "Carga"
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de la columna M en la hoja de "Carga"
  var datosM = hojaCarga.getRange("M18:M").getValues();
  var numRows = datosM.length;

  // Recorrer los datos
  for (var i = 0; i < numRows; i++) {
    var valorM = datosM[i][0];
    var valorO = "";

    // Determinar el valor de la columna O según el valor en la columna M
    switch (valorM) {
      case "TM Tractomula":
      case "TM2 Tractomula 2 ejes":
      case "TM3 Tractomula 3 ejes":
        valorO = "7,5";
        break;
      case "DT Dobletroque":
        valorO = "12";
        break;
      case "MM Minimula":
        valorO = "10";
        break;
      case "SC Sencillo":
        valorO = "14";
        break;
      case "TB Turbo":
        valorO = "15";
        break;
      default:
        valorO = "";
        break;
    }

    // Colocar el valor en la columna O de la fila actual
    hojaCarga.getRange(18 + i, 15).setValue(valorO);
  }
}
//Funcion para limpiar columnas STUV
function clearColumnsSTUV() {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener la hoja "Carga"
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");

  // Obtener el rango de las columnas S, T, U y V
  var rangoSTUV = hojaCarga.getRange("S18:V");

  // Limpiar el contenido de las celdas en el rango especificado
  rangoSTUV.clearContent();
}

//Funcion para obtener la ciudad y departamento, Origen,Destino y Distancia en Km
function originDestinRoute() {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener las hojas
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");
  var hojaKmTipoTransporte = archivoDestino.getSheetByName("Km-Tipo Transporte");

  // Obtener los datos de las columnas S en la hoja de "Carga"
  var datosCarga = hojaCarga.getRange("S18:S").getValues();
  //Logger.log("Datos de la columna S en 'Carga': " + JSON.stringify(datosCarga));
  var numRows = datosCarga.length;

  // Obtener los datos de la columna A en la hoja de "Km-Tipo Transporte"
  var datosKmTipoTransporte = hojaKmTipoTransporte.getRange("A2:A").getValues();
  //Logger.log("Datos de la columna A en 'Km-Tipo Transporte': " + JSON.stringify(datosKmTipoTransporte));

  // Recorrer los datos
  for (var i = 0; i < numRows; i++) {
    var valorS = datosCarga[i][0].toString().trim();
    //Logger.log("Valor de S en la fila " + (18 + i) + ": " + valorS);
    
    // Buscar el valor en la hoja de "Km-Tipo Transporte"
    var index = -1;
    for (var j = 0; j < datosKmTipoTransporte.length; j++) {
      var valorA = datosKmTipoTransporte[j][0].toString().trim();
      if (valorA === valorS) {
        index = j;
        break;
      }
    }
    //Logger.log("Índice encontrado: " + index);

    // Si se encuentra el valor, copiar los valores de las columnas B, C, D, E y F y pegarlos en la hoja de "Carga"
    if (index !== -1) {
      var valorB = hojaKmTipoTransporte.getRange(index + 2, 2).getValue(); // +2 para ajustar al índice base 1 y salto de encabezado
      var valorC = hojaKmTipoTransporte.getRange(index + 2, 3).getValue(); // +2 para ajustar al índice base 1 y salto de encabezado
      var valorD = hojaKmTipoTransporte.getRange(index + 2, 4).getValue(); // +2 para ajustar al índice base 1 y salto de encabezado
      var valorE = hojaKmTipoTransporte.getRange(index + 2, 5).getValue(); // +2 para ajustar al índice base 1 y salto de encabezado
      var valorF = hojaKmTipoTransporte.getRange(index + 2, 6).getValue(); // +2 para ajustar al índice base 1 y salto de encabezado
      //Logger.log("Valor B: " + valorB + ", Valor C: " + valorC + ", Valor D: " + valorD + ", Valor E: " + valorE + ", Valor F: " + valorF);
      
      hojaCarga.getRange(18 + i, 7).setValue(valorB); // Colocar en la columna G de la hoja de "Carga"
      hojaCarga.getRange(18 + i, 8).setValue(valorC); // Colocar en la columna H de la hoja de "Carga"
      hojaCarga.getRange(18 + i, 9).setValue(valorD); // Colocar en la columna I de la hoja de "Carga"
      hojaCarga.getRange(18 + i, 10).setValue(valorE); // Colocar en la columna J de la hoja de "Carga"
      hojaCarga.getRange(18 + i, 11).setValue(valorF); // Colocar en la columna K de la hoja de "Carga"
    }
  }
}

//Funcion para calcular el Rendimiento de Combustible por Viaje (KM/Galon)
//Operacion: Distancia recorrida por viaje/ Cantidad de Combustible por Viaje
function calculateFuelEfficiency() {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  
  // Abrir el archivo de destino y obtener la hoja "Carga"
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de las columnas K y O desde la fila 18 en adelante
  var datosDistancia = hojaCarga.getRange("K18:K").getValues(); // Distancia recorrida por viaje
  var datosCombustible = hojaCarga.getRange("O18:O").getValues(); // Cantidad de combustible por viaje
  var numRows = datosDistancia.length;

  // Crear un array para almacenar los resultados del rendimiento de combustible
  var resultados = [];

  // Recorrer los datos y realizar la división (KM/Galón)
  for (var i = 0; i < numRows; i++) {
    var distancia = datosDistancia[i][0];
    var combustible = datosCombustible[i][0];

    // Si alguno de los valores es nulo o vacío, dejar la celda vacía
    if (distancia === "" || combustible === "" || distancia === null || combustible === null) {
      resultados.push([""]);
    } else {
      var rendimiento = (combustible !== 0) ? distancia / combustible : 0; // Evitar división por cero
      resultados.push([rendimiento]);
    }
  }

  // Pegar los resultados en la columna P desde la fila 18 en adelante
  hojaCarga.getRange(18, 16, numRows, 1).setValues(resultados);
}
