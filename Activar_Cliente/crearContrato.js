function crearContrato() {
 
    var SSmaestroContratos = SpreadsheetApp.openById(contractmaestroId);
    var sheetContracts = SSmaestroContratos.getSheetByName("Datos");
    var contractLastRow = sheetContracts.getLastRow();
    var ContractLastColumn = sheetContracts.getLastColumn();
    var nit = sheetContracts.getRange(contractLastRow,6).getValue();

    var razonSocial = sheetContracts.getRange(contractLastRow,7).getValue();
    var repLegal = sheetContracts.getRange(contractLastRow,8).getValue();
    var numCedula = sheetContracts.getRange(contractLastRow,7).getValue();
    var ciudadCliente = sheetContracts.getRange(contractLastRow,9).getValue();
    var dirCliente = sheetContracts.getRange(contractLastRow,10).getValue();
    var valor = sheetContracts.getRange(contractLastRow,3).getValue();
    var valPesosCol = formatoColombiano(valor);
    var cot = sheetContracts.getRange(contractLastRow,2).getValue();
    var rut = sheetContracts.getRange(contractLastRow,12).getValue();
    var camaraDeComercio = sheetContracts.getRange(contractLastRow,13).getValue();
    var cedula = sheetContracts.getRange(contractLastRow,14).getValue();



    var parts = cot.split("-");
    var meses = 0;
    var plantilla = "";

    if(parts[3]== "7"){
    meses = 3;
    plantilla = contrato7estandares;
    numClien = parts[5]
    }
    else{
    meses = 12;
    plantilla = contratoSgsst;
    numClien = parts[3]
    }


    var numEmp = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","¿Cuántos trabajadores tiene actualmente directos?*");



    sheetContracts.getRange(contractLastRow,19).setValue(meses);
    sheetContracts.getRange(contractLastRow,17).setValue(numClien);
    
    let fechaInicio = sheetContracts.getRange(contractLastRow,4).getValue();
    let valorTotal = valor*meses
    let valtotpesosCol = formatoColombiano(valorTotal)
    let mesesL = numeroALetras(meses)
    let meseLe = mesesL.split(" ")
      
    //fecha de hoy
    var today = new Date();
    var yyyy = today.getFullYear();
    
   //Valores A letras
    var valorLetras = numeroALetras(valor, {
      plural: "PESOS",
      singular: "PESO",
      centPlural: "CENTAVOS",
      centSingular: "CENTAVO"
    });
  
    var valorTotalLetras = numeroALetras(valorTotal, {
      plural: "PESOS",
      singular: "PESO",
      centPlural: "CENTAVOS",
      centSingular: "CENTAVO"
    });
    
    var fecIni = new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), fechaInicio.getDate());
    var fechaFin = new Date(fechaInicio.getFullYear(), fechaInicio.getMonth() + meses, fechaInicio.getDate());
    
    // Obtener Data de archivo clientes-Cotizaciones
  
    
    
    var contractFolder = DriveApp.getFolderById(createContractFolder());
  
     
    //obtener plantilla contratos y crear copia
    var idPlantilla = DriveApp.getFileById(plantilla);
    var copy = idPlantilla.makeCopy("CONTRATO-"+cot,contractFolder);
    var urlContrato = copy.getUrl();
    sheetContracts.getRange(contractLastRow,ContractLastColumn).setValue(urlContrato);
    var docId = copy.getId();
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    var header = doc.getHeader();
  
    
  
    //Llenar contrato
   
    body.replaceText("{{repLegal}}", repLegal.toString().toUpperCase());
    body.replaceText("{{numCedula}}", numCedula);
    body.replaceText("{{razonSocial}}", razonSocial.toString().toUpperCase());
    header.replaceText("{{razonSocial}}", razonSocial.toString().toUpperCase());
    body.replaceText("{{nit}}", nit);
    body.replaceText("{{numEmp}}", numEmp);
    body.replaceText("{{ciudadCliente}}", ciudadCliente.toString().toUpperCase());
    body.replaceText("{{dirCliente}}", dirCliente);
    body.replaceText("{{anio}}", yyyy);
    body.replaceText("{{valorLetras}}", valorLetras);
    body.replaceText("{{valor}}", valPesosCol);
    body.replaceText("{{valorTotalLetras}}", valorTotalLetras);
    body.replaceText("{{valorTotal}}", valtotpesosCol);
    body.replaceText("{{fechaFin}}", convertirFecha(fechaFin));
    body.replaceText("{{numCotizacion}}", cot);
    body.replaceText("{{fechaActual}}", convertirFecha(today));
    body.replaceText("{{fechaInicio}}", convertirFecha(fecIni));
    body.replaceText("{{meses}}", meses);
    body.replaceText("{{mesesLetras}}", meseLe[0]);
  
  
    //Trasladar archivos adjuntos
  
    
  
    var contractFolderID = contractFolder.getId();
    var contractFolderUrl = contractFolder.getUrl(); 
    sheetContracts.getRange(contractLastRow,18).setValue(contractFolderID);  
  
    trasladarArchivo(contractFolderID, rut,"RUT-"+ razonSocial)
    trasladarArchivo(contractFolderID, camaraDeComercio,"Camara de Comercio "+ razonSocial )
    trasladarArchivo(contractFolderID, cedula,"Cedula Rep. Legal " + razonSocial)
    // formulario de aprobacion
  
    var form =preFilledForm(nit,razonSocial,cot,valor);
  
    // envio de mail para aprobacion
  
    
    sendEmail(cot,form,contractFolderUrl)
  
  }

function createContractFolder() {
  var folderId = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","clientFolderId");
  var folderName = "CONTRATOS";

  
  // Buscar la carpeta "contratos" en el folder padre
  var parentFolder = DriveApp.getFolderById(folderId);
  var existingFolders = parentFolder.getFoldersByName(folderName);
  
  
  // Si la carpeta "contratos" ya existe, devolver el ID de esa carpeta
  if (existingFolders.hasNext()) {
    var existingFolder = existingFolders.next();
    var existingFolderId = existingFolder.getId();
    Logger.log("Carpeta 'contratos' encontrada. ID: " + existingFolderId);
    return existingFolderId;
  }
  
  // Si la carpeta "contratos" no existe, crear una nueva y devolver el ID de la carpeta recién creada
  var newFolder = parentFolder.createFolder(folderName);
  var newFolderId = newFolder.getId();
  Logger.log("Nueva carpeta 'contratos' creada. ID: " + newFolderId);
  return newFolderId;

}

function sendEmail(numCot,link,carpeta) {
  var email = "alejandrosaga61@gmail.com"
  var email2 = "bitsaga2804@gmail.com"
  var firstName = "Nancy";
  var subject = "Revisar contrato: " + numCot+2;
  var body = '<p>Hola <strong>'+ firstName +'</strong>, Tenemos un nuevo cliente!! La empresa <strong>'+razonSocial+'</strong> ha decidido aceptar la cotizacion '+numCot+'.</p> <p>A el borrador de contrato y la documentacion adjunta por el cliente para revision se encuentra en la siguiente carpeta'+carpeta+' y el link correspondiente al formulario de aprobacion.'+link+'<p>Una vez envies el formulario aceptando el contrato, este será enviado al cliente en PDF 2. </p>'
  MailApp.sendEmail({
      to: email,
      cc: email2,
      subject: subject,
      htmlBody: body
      
    }); 
}

function trasladarArchivo(idCarpetaDestino, urlArchivo, nuevoNombre) {
  var carpetaDestino = DriveApp.getFolderById(idCarpetaDestino);
  var archivo = DriveApp.getFileById(getIdFromUrl(urlArchivo));
  
  var copiaArchivo = archivo.makeCopy(nuevoNombre, carpetaDestino);
  
  // Opcional: Eliminar el archivo de su ubicación original
  archivo.setTrashed(true);
  
  var nuevaUrl = copiaArchivo.getUrl();
  
  return nuevaUrl;
}

// Función auxiliar para extraer el ID de un enlace de Google Drive
function getIdFromUrl(url) {
  var id = "";
  var match = url.match(/[-\w]{25,}/);
  if (match) {
    id = match[0];
  }
  return id;
}
function preFilledForm(id,nombre,contrato,valor){

 
  
    //Form aprobacion pre-llenado
    
    //var idEmpresa = id.replace(/ /g, '+');
    var contratoName = contrato.replace(/ /g, '+');
    var companyName = nombre.replace(/ /g, '+');
  
    
    
    
    var prefilledForm= "https://docs.google.com/forms/d/e/1FAIpQLScnKWRXGB7XND1zK7msYKoLYEqP1eL-LK7zKEjjW9IFnbMw7A/viewform?usp=pp_url&entry.2087970223="+id+"&entry.653991903="+companyName+"&entry.1862569191="+contratoName+"&entry.120278530="+valor;
  
    return prefilledForm;
  }
  
  
  
  