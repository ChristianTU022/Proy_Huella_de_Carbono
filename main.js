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
  // Conectar Sheets a AppScript
  const sheetId = '19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA';
  const sheet = SpreadsheetApp.openById(sheetId);
  
  // Retornar hojas específicas
  return {
      sheet,
      p_Carga: sheet.getSheetByName('Carga'),
      p_Transportadoras: sheet.getSheetByName('TRANSPORTADORAS'),
      p_Km_TipoTransporte: sheet.getSheetByName('Km-Tipo Transporte')
  };
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
  // Datos del Archivo Origen
  var fechaActual = new Date();
  var mesAnterior = fetchLastMonth(); // Obtener el mes anterior
  var currentYear = fechaActual.getFullYear(); // Obtener el año actual
  var nombreArchivoOrigen = "[GSheets-]VT12 " + mesAnterior + " " + currentYear;

  // Datos del Archivo Destino
  var filaInicioDestino = 18;

  // Obtener las hojas específicas
  const { p_Carga } = conectionSheets();

  var archivosOrigen = DriveApp.getFilesByName(nombreArchivoOrigen);
  if (archivosOrigen.hasNext()) {
      var archivoOrigen = archivosOrigen.next();
      var hojaOrigen = SpreadsheetApp.openById(archivoOrigen.getId()).getSheetByName("Hoja1");
      var datosOrigen = hojaOrigen.getDataRange().getValues().slice(1); // Saltar la primera fila

      // Filtrar los datos de la columna B que no comiencen por "580"
      var datosFiltrados = datosOrigen.filter(function (fila) {
          return !fila[1] || fila[1].toString().indexOf("580") !== 0;
      });

      // Calcular la cantidad de filas de datos a copiar
      var numRows = datosFiltrados.length;

      // Pegar los datos en la hoja de destino
      p_Carga.getRange(filaInicioDestino, 2, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[0]]; }));

      // Agregar "Pastas" en la columna A y "Seco" en la columna F
      p_Carga.getRange(filaInicioDestino, 1, numRows, 1).setValue("Pastas");
      p_Carga.getRange(filaInicioDestino, 6, numRows, 1).setValue("Seco");

      // Verificar las categorías y colocar "Primario" o "Secundario" en la columna C
      p_Carga.getRange(filaInicioDestino, 3, numRows, 1).setValues(datosFiltrados.map(function (fila) {
          if (["ZP01", "ZP02", "ZP07"].includes(fila[12])) {
              return ["Primario"];
          } else if (["ZP03", "ZP04", "ZP05", "ZP06", "ZP08"].includes(fila[12])) {
              return ["Secundario"];
          } else {
              return [""];
          }
      }));

      // Copiar datos adicionales del archivo de origen al archivo de destino
      p_Carga.getRange(filaInicioDestino, 17, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[28]]; }));
      p_Carga.getRange(filaInicioDestino, 4, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[5]]; }));
      p_Carga.getRange(filaInicioDestino, 5, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[23]]; }));
      p_Carga.getRange(filaInicioDestino, 20, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[11]]; }));
      p_Carga.getRange(filaInicioDestino, 19, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[1]]; }));
      p_Carga.getRange(filaInicioDestino, 12, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[47]]; }));
      p_Carga.getRange(filaInicioDestino, 21, numRows, 1).setValues(datosFiltrados.map(function (fila) { return [fila[7]]; })); // Columna U del archivo origen

      // Verificar si hay datos en las columnas O, P o Q y colocar 1 o 0 en la columna V del destino
      p_Carga.getRange(filaInicioDestino, 22, numRows, 1).setValues(datosFiltrados.map(function (fila) {
          return [(fila[14] || fila[15] || fila[16] || fila[17] || fila[18] || fila[19] || fila[20] || fila[21]) ? 1 : 0];
      }));
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
function removeSpecificRows() {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener la hoja
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaDestino = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de la hoja
  var datos = hojaDestino.getDataRange().getValues();

  // Crear un arreglo para almacenar las filas a eliminar
  var filasEliminar = [];

  // Recorrer los datos para identificar las filas a eliminar
  datos.forEach(function(fila, index) {
    // Eliminar filas que cumplan con las condiciones
    if (
      fila[1].toString().indexOf("580") === 0 || // Comienza con "580" en la columna B
      fila[18].toString() === "PINTER" || // Tiene "PINTER" en la columna S
      fila[19].toString() === "DR11" || fila[19].toString() === "DR15" || // Tiene "DR11" o "DR15" en la columna T
      fila[21] === 0 // Tiene 0 en la columna V
    ) {
      filasEliminar.push(index + 1); // Se agrega el índice de la fila para eliminarla
    }
  });

  // Eliminar las filas del arreglo en orden inverso para evitar problemas de índices
  filasEliminar.reverse().forEach(function(indice) {
    hojaDestino.deleteRow(indice);
  });
}

//Funcion relacionada al Boton con el fin de llamar otras funciones
function completeTableFields () {
  //Llamar a la función getNumberOfRowsWithData que obtiene el numero total de filas
  var numRows = getNumberOfRowsWithData();
  //Llamar a la función typeOfBelongingAndFuel "Tipo de Flota o Pertenencia (Dice si es Propio o Contratado)"
  typeOfBelongingAndFuel(numRows);
  // Llamar a la función findTypeTransportation para determinar el tipo de transporte que se uso. (SC Sencillo,TB Turbo, TM2 Tractomula 2 ejes, DT Dobletroque, etc)
  findTypeTransportation(numRows);
  //Llamar a la funcion amountOfFuelPerTrip para determinar la cantidad de Combustible por Viaje (gal/viaje)
  amountOfFuelPerTrip(numRows);
  //Llama a la funcion originDestinRoute, la cual encuentra la ciudad y origen de destino
  originDestinRoute(numRows);
  //Llama a la funcion ,la cual calcula el Rendimiento por Viaje (Km/Gal)
  calculateFuelEfficiency(numRows);
  //Llamar a la funcion searchNITTransporter, la cual busca el NIT de la Transportadora
  searchNITTransporter(numRows);
  // //Llamar a la funcion clearColumnsSTUV, la cual despues de llenar todo borra los datos temporales de la columna S, T, U y V
  clearColumnsSTUV();
}

//Funcion para saber el numero de filas para romper los ciclos al buscar la informacion
function getNumberOfRowsWithData() {
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaDestino = archivoDestino.getSheetByName("Carga");

  // Obtener el rango de la columna B desde la fila 18
  var datosB = hojaDestino.getRange("B18:B").getValues();
  
  // Calcular el número de filas con datos en la columna B
  var numRows = datosB.filter(String).length;
  Logger.log("Número de filas con datos: " + numRows);

  return numRows;
}

//Funcion para saber si el transporte es propio o contratado y el tipo de Combustible
function typeOfBelongingAndFuel(numRows) {
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  
  // Abrir el archivo de destino
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaDestino = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de la columna Q desde la fila 18 hasta el número de filas con datos
  var datosQ = hojaDestino.getRange(18, 17, numRows, 1).getValues(); // Columna Q
  
  // Crear arrays para almacenar los resultados
  var resultadosR = [];
  var resultadosN = [];

  // Recorrer los datos
  for (var i = 0; i < numRows; i++) {
    var valorQ = datosQ[i][0];
    var valorR = valorQ === "JPO336" ? "Propio" : valorQ ? "Tercerizado" : "";
    var tipoCombustible = valorR === "Propio" ? "Gas Natural" : valorR === "Tercerizado" ? "Diésel" : "";

    // Almacenar los resultados en los arrays
    resultadosR.push([valorR]);
    resultadosN.push([tipoCombustible]);
  }

  // Colocar el tipo de pertenencia en la columna R desde la fila 18 en adelante
  hojaDestino.getRange(18, 18, numRows, 1).setValues(resultadosR); // Columna R (índice 18)
  
  // Colocar el tipo de combustible en la columna N desde la fila 18 en adelante
  hojaDestino.getRange(18, 14, numRows, 1).setValues(resultadosN); // Columna N (índice 14)
}

//Funcion que con los pesos Dice que tipo de Vehiculo es "Sencillo, Turbo, TM 2 Ejes, TM 3 Ejes, etc"
function findTypeTransportation(numRows) {
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaDestino = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de las columnas L y U desde la fila 18 hasta el número de filas con datos
  var datos = hojaDestino.getRange(18, 12, numRows, 10).getValues(); // Columnas L hasta U
  var resultados = [];

  for (var i = 0; i < numRows; i++) {
    var valorL = datos[i][0];  // Columna L
    var textoU = datos[i][9].toUpperCase();  // Columna U

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

    // Verificar si el texto contiene "MINI"
    if (textoU.indexOf("MINI") !== -1) {
      tipoVehiculo = "MM Minimula";
    }

    // Almacenar el tipo de vehículo en el array de resultados
    resultados.push([tipoVehiculo]);
  }

  // Colocar el tipo de vehículo en la columna M desde la fila 18 en adelante
  hojaDestino.getRange(18, 13, numRows, 1).setValues(resultados);
}


// Esta funcion busca el NIT de la transportadora por su nombre
function searchNITTransporter(numRows) {
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener las hojas
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");
  var hojaTransportadoras = archivoDestino.getSheetByName("TRANSPORTADORAS");

  // Obtener los datos de las columnas D en la hoja de "Carga"
  var datosCarga = hojaCarga.getRange(18, 4, numRows, 1).getValues();

  // Obtener los datos de las columnas B y C en la hoja de "TRANSPORTADORAS"
  var datosTransportadoras = hojaTransportadoras.getRange(1, 2, hojaTransportadoras.getLastRow(), 2).getValues();

  // Crear un objeto para mapear NITs a nombres de transportadoras
  var mapaTransportadoras = {};
  datosTransportadoras.forEach(function(fila) {
    mapaTransportadoras[fila[0]] = fila[1];
  });

  // Recorrer los datos de carga y asignar el nombre de la transportadora si se encuentra el NIT
  datosCarga.forEach(function(fila, index) {
    var valorD = fila[0];
    if (mapaTransportadoras.hasOwnProperty(valorD)) {
      var valorC = mapaTransportadoras[valorD];
      hojaCarga.getRange(18 + index, 4).setValue(valorC);
    }
  });
}



// Esta funcion tiene el fin de decir cual es la cantidad de combustible por viaje (gal/viaje), con ayuda del Tipo de Vehiculo
function amountOfFuelPerTrip(numRows) {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener la hoja "Carga"
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de la columna M desde la fila 18 hasta el número de filas con datos
  var datosM = hojaCarga.getRange(18, 13, numRows, 1).getValues(); // Columna M

  // Crear un array para almacenar los resultados
  var resultadosO = [];

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

    // Almacenar el resultado en el array
    resultadosO.push([valorO]);
  }

  // Colocar los valores en la columna O desde la fila 18 en adelante
  hojaCarga.getRange(18, 15, numRows, 1).setValues(resultadosO); // Columna O (índice 15)
}

//Funcion para limpiar columnas STUV
function clearColumnsSTUV() {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener la hoja "Carga"
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");

  // Limpiar el contenido de las columnas S, T, U y V en la hoja "Carga"
  hojaCarga.getRange("S18:V").clearContent();
}


//Funcion para obtener la ciudad y departamento, Origen,Destino y Distancia en Km
function originDestinRoute(numRows) {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";

  // Abrir el archivo de destino y obtener las hojas
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");
  var hojaKmTipoTransporte = archivoDestino.getSheetByName("Km-Tipo Transporte");

  // Obtener los datos de las columnas S en la hoja de "Carga" desde la fila 18 hasta el número de filas con datos
  var datosCarga = hojaCarga.getRange(18, 19, numRows, 1).getValues().flat();

  // Obtener los datos de la hoja "Km-Tipo Transporte"
  var datosKmTipoTransporte = hojaKmTipoTransporte.getDataRange().getValues();

  // Crear un objeto para mapear los valores de la columna A a las columnas B, C, D, E y F
  var mapaKmTipoTransporte = {};
  for (var i = 1; i < datosKmTipoTransporte.length; i++) { // Empezamos desde 1 para saltar el encabezado
    var valorA = datosKmTipoTransporte[i][0].toString().trim();
    mapaKmTipoTransporte[valorA] = datosKmTipoTransporte[i].slice(1, 6); // Obtener B, C, D, E, F
  }

  // Crear un array para almacenar los resultados
  var resultados = [];

  // Recorrer los datos y buscar en el mapa
  for (var i = 0; i < numRows; i++) {
    var valorS = datosCarga[i].toString().trim();
    var valores = mapaKmTipoTransporte[valorS] || ["", "", "", "", ""];
    resultados.push(valores);
  }

  // Escribir los resultados en las columnas G, H, I, J y K de la hoja "Carga"
  hojaCarga.getRange(18, 7, numRows, 5).setValues(resultados);
}



//Funcion para calcular el Rendimiento de Combustible por Viaje (KM/Galon)
//Operacion: Distancia recorrida por viaje/ Cantidad de Combustible por Viaje
function calculateFuelEfficiency(numRows) {
  // ID del archivo en el que se trabajarán los datos
  var idArchivo = "19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA";
  
  // Abrir el archivo de destino y obtener la hoja "Carga"
  var archivoDestino = SpreadsheetApp.openById(idArchivo);
  var hojaCarga = archivoDestino.getSheetByName("Carga");

  // Obtener los datos de las columnas K y O desde la fila 18 hasta el número de filas con datos
  var rangoDatos = hojaCarga.getRange(18, 11, numRows, 2).getValues(); // Obtiene K y O juntos
  
  // Crear un array para almacenar los resultados del rendimiento de combustible
  var resultados = rangoDatos.map(function(fila) {
    var distancia = fila[0];
    var combustible = fila[1];
    // Si alguno de los valores es nulo o vacío, dejar la celda vacía
    if (distancia === "" || combustible === "" || distancia === null || combustible === null) {
      return [""];
    } else {
      var rendimiento = (combustible !== 0) ? distancia / combustible : 0; // Evitar división por cero
      return [rendimiento];
    }
  });

  // Pegar los resultados en la columna P desde la fila 18 en adelante
  hojaCarga.getRange(18, 16, numRows, 1).setValues(resultados);
}
