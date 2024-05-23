function crearContrato() {
 
    let SSmaestroContratos = SpreadsheetApp.openById(contractmaestroId);
    let sheetContracts = SSmaestroContratos.getSheetByName("Datos");
    let contractLastRow = sheetContracts.getLastRow();
    let ContractLastColumn = sheetContracts.getLastColumn();
    let nit = sheetContracts.getRange(contractLastRow,6).getValue();

    let razonSocial = sheetContracts.getRange(contractLastRow,7).getValue();
    let repLegal = sheetContracts.getRange(contractLastRow,8).getValue();
    let numCedula = sheetContracts.getRange(contractLastRow,9).getValue();
    let ciudadCliente = sheetContracts.getRange(contractLastRow,10).getValue();
    let dirCliente = sheetContracts.getRange(contractLastRow,11).getValue();
    let valor = sheetContracts.getRange(contractLastRow,3).getValue();
    let valPesosCol = formatoColombiano(valor);
    let cot = sheetContracts.getRange(contractLastRow,2).getValue();
    let rut = sheetContracts.getRange(contractLastRow,12).getValue();
    let camaraDeComercio = sheetContracts.getRange(contractLastRow,13).getValue();
    let cedula = sheetContracts.getRange(contractLastRow,14).getValue();



    let parts = cot.split("-");
    let meses = 0;
    let plantilla = "";
    let numClien = "";
    let valPer = 0;
    let valAnti = 0;
    let numTra = 0;

   

    if (parts[3] == "7") {
      meses = 3;
      plantilla = contrato7estandares;
      numClien = parts[5];
    } else if (parts[2] == "PSI") {
      meses = 3;
      plantilla = contratoPSI;
      numClien = parts[3];
      valPer = searchValues(batPsiServiceId, cot, "Aplicacion Bateria riesgo psicosocial", "slideName", "ValorPersona");
      valAnti = searchValues(batPsiServiceId, cot, "Aplicacion Bateria riesgo psicosocial", "slideName", "valAnti");
      numTra = searchValues(maestroCotId, numClien, "Datos", "Codigo Cliente", "Por favor indique la cantidad de trabajadores que deben aplicar para la bateria de riesgo Psicosocial.");
    } else {
      meses = 12;
      plantilla = contratoSgsst;
      numClien = parts[3];
    }


    let numEmp = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","¿Cuántos trabajadores tiene actualmente directos?*");



    sheetContracts.getRange(contractLastRow,19).setValue(meses);
    sheetContracts.getRange(contractLastRow,17).setValue(numClien);
    
    let fechaInicio = sheetContracts.getRange(contractLastRow,4).getValue();
    let valorTotal = valor*meses
    let valPerPesos = formatoColombiano(valPer)
    let valAntiPesos = formatoColombiano(valAnti)
    let valtotpesosCol = formatoColombiano(valorTotal)
    let mesesL = numeroALetras(meses)
    let meseLe = mesesL.split(" ")
      
    //fecha de hoy
    let today = new Date();
    let yyyy = today.getFullYear();
    
   //Valores A letras
   let valPerLetras = numeroALetras(valPer, {
    plural: "PESOS",
    singular: "PESO",
    centPlural: "CENTAVOS",
    centSingular: "CENTAVO"
  });

  let valAntiLetras = numeroALetras(valAnti, {
    plural: "PESOS",
    singular: "PESO",
    centPlural: "CENTAVOS",
    centSingular: "CENTAVO"
  });

    let valorLetras = numeroALetras(valor, {
      plural: "PESOS",
      singular: "PESO",
      centPlural: "CENTAVOS",
      centSingular: "CENTAVO"
    });
  
    let valorTotalLetras = numeroALetras(valorTotal, {
      plural: "PESOS",
      singular: "PESO",
      centPlural: "CENTAVOS",
      centSingular: "CENTAVO"
    });
    
    let fecIni = new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), fechaInicio.getDate());
    let fechaFin = new Date(fechaInicio.getFullYear(), fechaInicio.getMonth() + meses, fechaInicio.getDate());
    
    // Obtener Data de archivo clientes-Cotizaciones
  
    
    
    let contractFolder = DriveApp.getFolderById(createContractFolder(numClien));
  
     
    //obtener plantilla contratos y crear copia
    let idPlantilla = DriveApp.getFileById(plantilla);
    let copy = idPlantilla.makeCopy("CONTRATO-"+cot,contractFolder);
    let urlContrato = copy.getUrl();
    sheetContracts.getRange(contractLastRow,ContractLastColumn).setValue(urlContrato);
    let docId = copy.getId();
    let doc = DocumentApp.openById(docId);
    let body = doc.getBody();
    let header = doc.getHeader();
  
    
  
    //Llenar contrato
   
    body.replaceText("{{repLegal}}", repLegal.toString().toUpperCase());
    body.replaceText("{{numCedula}}", numCedula);
    body.replaceText("{{razonSocial}}", razonSocial.toString().toUpperCase());
    header.replaceText("{{razonSocial}}", razonSocial.toString().toUpperCase());
    body.replaceText("{{nit}}", nit);
    body.replaceText("{{numEmp}}", numEmp);
    body.replaceText("{{numTra}}", numTra);
    body.replaceText("{{ciudadCliente}}", ciudadCliente.toString().toUpperCase());
    body.replaceText("{{dirCliente}}", dirCliente);
    body.replaceText("{{anio}}", yyyy);
    body.replaceText("{{valPer}}", valPerPesos);
    body.replaceText("{{valPerletras}}", valPerLetras);
    body.replaceText("{{valAntiPesos}}", valAntiPesos);
    body.replaceText("{{valAntiLetras}}", valAntiLetras);
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
  
    
  
    let contractFolderID = contractFolder.getId();
    let contractFolderUrl = contractFolder.getUrl(); 
    sheetContracts.getRange(contractLastRow,18).setValue(contractFolderID);  
  
    trasladarArchivo(contractFolderID, rut,"RUT-"+ razonSocial)
    trasladarArchivo(contractFolderID, camaraDeComercio,"Camara de Comercio "+ razonSocial )
    trasladarArchivo(contractFolderID, cedula,"Cedula Rep. Legal " + razonSocial)
    // formulario de aprobacion
  
    let contratoName = cot.replace(/ /g, '+');
    let companyName = razonSocial.replace(/ /g, '+');  
    let form= "https://docs.google.com/forms/d/e/"+approveContractForm+"/viewform?usp=pp_url&entry.2087970223="+nit+"&entry.653991903="+companyName+"&entry.1862569191="+contratoName+"&entry.120278530="+valor;
  
  
    // envio de mail para aprobacion
    let firstName = "Nancy";
    let subject = "Revisar contrato: " + cot;
    let emailBody = '<p>Hola <strong>'+ firstName +'</strong>, Tenemos un nuevo cliente!! La empresa <strong>'+razonSocial+'</strong> ha decidido aceptar la cotizacion '+cot+'.</p> <p>A el borrador de contrato y la documentacion adjunta por el cliente para revision se encuentra en la siguiente carpeta'+contractFolderUrl+' y el link correspondiente al formulario de aprobacion.'+form+'<p>Una vez envies el formulario aceptando el contrato, este será enviado al cliente en PDF 2. </p>'
    MailApp.sendEmail({
        to: personycaEmail1,
        cc: personycaEmail2,
        subject: subject,
        htmlBody: emailBody
        
        }); 
       
  }

function createContractFolder(numClien) {
  let folderId = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","clientFolderId");
  let folderName = "CONTRATOS";

  
  // Buscar la carpeta "contratos" en el folder padre
  let parentFolder = DriveApp.getFolderById(folderId);
  let existingFolders = parentFolder.getFoldersByName(folderName);
  
  
  // Si la carpeta "contratos" ya existe, devolver el ID de esa carpeta
  if (existingFolders.hasNext()) {
    let existingFolder = existingFolders.next();
    let existingFolderId = existingFolder.getId();
    Logger.log("Carpeta 'contratos' encontrada. ID: " + existingFolderId);
    return existingFolderId;
  }
  
  // Si la carpeta "contratos" no existe, crear una nueva y devolver el ID de la carpeta recién creada
  let newFolder = parentFolder.createFolder(folderName);
  let newFolderId = newFolder.getId();
  Logger.log("Nueva carpeta 'contratos' creada. ID: " + newFolderId);
  return newFolderId;

}

function trasladarArchivo(idCarpetaDestino, urlArchivo, nuevoNombre) {
  let carpetaDestino = DriveApp.getFolderById(idCarpetaDestino);
  let archivo = DriveApp.getFileById(getIdFromUrl(urlArchivo));
  
  let copiaArchivo = archivo.makeCopy(nuevoNombre, carpetaDestino);
  
  // Opcional: Eliminar el archivo de su ubicación original
  archivo.setTrashed(true);
  
  let nuevaUrl = copiaArchivo.getUrl();
  
  return nuevaUrl;
}

// Función auxiliar para extraer el ID de un enlace de Google Drive
function getIdFromUrl(url) {
  let id = "";
  let match = url.match(/[-\w]{25,}/);
  if (match) {
    id = match[0];
  }
  return id;
}

  
  
  