function doGet(){
    return HtmlService.createHtmlOutputFromFile('index') 
                      .setSandboxMode(HtmlService.SandboxMode.IFRAME) //Cambio #1
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); //Cambio #2
}


function importLocal () {
    let container = '<style>iframe{width: 100%; height:320px; border: none;}</style>'+
    '<iframe src="https://script.google.com/macros/s/AKfycbwGStrBPwh_NZSEDNhXes8j6YE-rmJJjxovJbd63v17qQKPSI92BUrThzhb3JfPz3zblQ/exec"></iframe>';

    let html = HtmlService.createHtmlOutput(container).setHeight(340).setWidth(550);
    SpreadsheetApp.getUi().showModalDialog(html, '🔎 Buscar archivo de Excel a Convertir') //.showSidebar(html);
}

function processForm(formObject) {
    let formBlob = formObject.myFile;
    //--------------- Inicia Adapcion ---------------
    // Ruta de Carpeta: https://drive.google.com/drive/u/0/folders/15gsKyGFx-+
    let carpetaAplicativo = DriveApp.getFileById("15gsKyGFx-UcK8TQXNdrfGtEWaSWPI9gK"); //Cambio #3
    let driveFile = carpetaAplicativo.createFile(formBlob); //Cambio #4

    let archivo = deExcelaGSheets(driveFile.getName()); //Convertimos con la funcion Aux y obtenemos el Objeto (#5)
    let nombreGSheet = archivo.name; // Obtenemos el nombre del archivo (ver funcion Auxiliar)(#6)
    let url = archivo.url; //Obtenemos la url del archivo (#7)
    let id = archivo.id; //Obtenemos su id todo desde el objeto "archivo" que nos da la funcion Auxiliar (#8)

    let ss = SpreadsheetApp.openById("19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA"); //Cambio #9. -Nuevo Archivo
    let hoja = ss.getSheetByName("Hoja_Url_Conversion"); //Cambio #10
    let marcaTemp = Utilities.formatDate(new Date(), "GWT", "dd-MM-yyyy HH:mm:ss"); //Cambio #11
    hoja.appendRow([marcaTemp, nombreGSheet, url, id]); //Cambio #12
    //--------------- Termina Adapcion ---------------
    return url; // Regresa la Url que se pasara como parametro a la web App para mostrar el enlace. (#13)
}

function deExcelaGSheets(name){
    let files = DriveApp.getFilesByName(name);
    let excelFile = null;
    if(files.hasNext())
        excelFile = files.next();
    else
        return null;
    let blob = excelFile.getBlob();

    let config = {
        title: "[Sheet-] " + excelFile.getName(),
        parents: [{id: excelFile.getParents().next().getId()}],
        // ruta:{ link: excelFile.getUrl()},
        mimeType: MimeType.GOOGLE_SHEETS
    };
    let spreadsheet = APIDrive.Files.insert(config, blob); // En esta linea se cambia de Excel a Sheets

    let archivo={
        name: spreadsheet.title,
        url: spreadsheet.alternateLink,
        id: spreadsheet.id
    }
    return archivo;
}
