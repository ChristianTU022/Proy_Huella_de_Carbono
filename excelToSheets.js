function importLocal(){
    var contenedor = '<style>iframe{width:100%;height:320px;border:none;}</style>'+
    '<iframe src="https://script.google.com/macros/s/AKfycbyFTl9cE2VPlhINDhH5ViA4MtkDExQYOuCjlyqVVx-fHv59i-odWm8DKp7nRcXu_Hq8Pw/exec"></iframe>';
    
    var html = HtmlService.createHtmlOutput(contenedor).setHeight(340).setWidth(550);
    SpreadsheetApp.getUi().showModalDialog(html, '🔎 Buscar archivo de Excel a Convertir') //.showSideBar(html);
}

function doGet(){
    return HtmlService.createHtmlOutputFromFile('index')
                      .setSandboxMode(HtmlService.SandboxMode.IFRAME)//Cambio #1
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);//Cambio #2
}
  
function processForm(formObject) {
    var formBlob = formObject.myFile;
    //------------ inicia adaptacion ----------------------------------------------
    //ver resultado en folder: https://drive.google.com/drive/folders/1_Xkb_TBY63MI1fMdNUQx8zM4jk8jEwyN
    var carpetaAplicativo = DriveApp.getFolderById("1_Xkb_TBY63MI1fMdNUQx8zM4jk8jEwyN");//Cambio #3
    var driveFile = carpetaAplicativo.createFile(formBlob);//Cambio #4

    var archivo = deExcelaGShets(driveFile.getName());//convertimos con la fun aux y obtenemos el objeto(#5)
    var nombreGSheet = archivo.nombre;//obtenemos el nombre del archivo (ver funcion auxiliar)(#6)
    var url = archivo.url;//obtenemos la url del archivo (#7)
    var id = archivo.id;//obtenemos su id todo desde el objeto "archivo" que nos da la fun auxiliar (#8)
  
    var ss = SpreadsheetApp.openById("19YHD7oJYoms0juBEp52rq4ljuqMucvR7gU-ZQd-ZCOA");//Cambio #9
    var hoja = ss.getSheetByName("Hoja_Url_Conversion");//Cambio #10
    var marcaTemp = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy HH:mm:ss");//Cambio #11
    hoja.appendRow([marcaTemp, nombreGSheet,url,id]);//Cambio #12
    return url; //regresa la url que se pasará como parametro a la web app para mostrar el enlace.(#13)
    //------------ termina adaptación ---------------------------------------------
}

function deExcelaGShets (nombre){
    var files = DriveApp.getFilesByName(nombre);
    var excelFile = null;
    if (files.hasNext())
        excelFile = files.next();
    else
        return null;
    var blob = excelFile.getBlob();

    var config = {
        title: "[GSheets-]" + excelFile.getName(),
        parents: [{id: excelFile.getParents().next().getId()}],
        MimeType: MimeType.GOOGLE_SHEETS
    };
    var spreadsheet = Drive.Files.insert(config, blob, {convert:true});

    var archivo={
        nombre: spreadsheet.title,
        url: spreadsheet.alternateLink,
        id: spreadsheet.id
    }
    
    return archivo;
}
  
