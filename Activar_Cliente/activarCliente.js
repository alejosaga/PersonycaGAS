async function activarCliente() {
  
    let ssmasterContractApprove = SpreadsheetApp.openById(contractApproveId)
    let sheetAprContr = ssmasterContractApprove.getSheetByName("Datos");
    let lastRowContApr = sheetAprContr.getLastRow();
    let contrato = sheetAprContr.getRange(lastRowContApr,3).getValue();
    let nombreRazonSocial = sheetAprContr.getRange(lastRowContApr,6).getValue();
    /*
    var parts = contrato.split("-");
    
    if(parts[3]== "7"){
        numClien = parts[5]
    }
    else{
      
      numClien = parts[3]
    }
    //let email1 = 'autopesonyca@gmail.com';
    let email2 = 'autopesonyca@gmail.com';
    
    var servicio = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","Servicios de interes")
  
    //2.Agregar fecha de inicio del contrato en la columna 7
      var fecIni = searchValues(contractmaestroId,numClien,"Datos","Codigo Cliente","Fecha de inicio");
      sheetAprContr.getRange(lastRowContApr,7).setValue(fecIni);
  
    //3.Buscar contrato en la carpeta del Cliente en "Prospectos".
  
      var folderId = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","clientFolderId");
      var servicio = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","Servicios de interes");
      sheetAprContr.getRange(lastRowContApr,10).setValue(servicio);
  
    //4.Convertir Contrato a PDF y dejarlo en la carpeta del cliente prospecto.
  
      var contratoUrl = searchValues(contractmaestroId,numClien,"Datos","Codigo Cliente","urlContrato");
      var contractFolderId = searchValues(contractmaestroId,numClien,"Datos","Codigo Cliente","Carpeta Contratos");
      var contratoId = getIdFromUrl(contratoUrl);
      var archivoPDFId = convertirDocAPDF(contratoId, contractFolderId);  
  
     //5.Trasladar carpeta del cliente desde "Prospectos" a "Activos".
  
      var newfolderId = trasladarCarpeta(folderId, folderClienteActivoId);
      sheetAprContr.getRange(lastRowContApr,8).setValue(newfolderId);
  
  
    //6.Enviar contrato firmado por Personyca al Cliente mediante Email.
  
      var email = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","Dirección de correo electrónico");
      var subject = " Contrato adjunto - Confirmar recepción y firma: " + contrato;
      var archivoAdjunto = DriveApp.getFileById(archivoPDFId);
      //sendEmail(subject,toClient,email,archivoAdjunto)
      let nombreCliente = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","Nombres y apellidos de la persona contacto")
      let dataToSend = '<p>Estimado(a),</p> <strong>'+ nombreCliente +'</strong>,</p><p>Adjunto encontrará el contrato: ' +contrato+ ' que hemos preparado para usted. Por favor, revise el contrato detenidamente y, si está de acuerdo con los términos, le agradeceríamos que lo firmara y nos lo devolviera por correo electrónico.</p><p>Si tiene alguna pregunta o inquietud con respecto al contrato, no dude en comunicarse con nosotros. Estaremos encantados de brindarle cualquier aclaración adicional que necesite.</p><p>Una vez recibido el contrato firmado, procederemos de inmediato con los trámites necesarios para avanzar en el proceso. Esperamos con entusiasmo comenzar esta relación de trabajo y colaborar juntos.</p><p>Adjunto encontrará el contrato firmado en formato PDF. Si prefiere recibirlo en otro formato o necesita asistencia adicional, por favor háganoslo saber.</p><p>Agradecemos su atención y cooperación en este asunto. Quedamos a su disposición para cualquier consulta adicional.</p><p>Saludos cordiales,</p><p>Nancy Camacho<br>Gerente<br>Personyca SAS<br>gerenciapersonyca@gmail.com<br>3165549102</p>'; 

      var body = dataToSend;
      MailApp.sendEmail({
          to: email,
          cc: email2,
          subject: subject,
          htmlBody: body,
          attachments: [archivoAdjunto.getAs(MimeType.PDF)]
        }); 
*/
    //1.Crear plan de Trabajo (Validar la mejor opcion entre crearlo desde una plantilla de Sheets o crearlo directamente en Click Up mediante la API).
    
        
    // Obtener el ID del espacio, ya sea existente o creado
        var spaceId = await createSpaceAndGetId(nombreRazonSocial);

    // Manejar el resultado
    if (spaceId) {
        console.log(`El espacio "${nombreRazonSocial}" tiene el ID: ${spaceId}`)};

            
    
   
    

      
    sheetAprContr.getRange(lastRowContApr,9).setValue(result);
    var spaceId = sheetAprContr.getRange(lastRowContApr,9).getValue();
    console.log(clickupfolder)
    
  
      
  
    //7.Guardar el plan de trabajo en carpeta del Cliente en caso de que se cree mediante Sheets.
    //8.Enviar Email a gerente de Personyca y consultores vinculados al proyecto, informando sobre la activacion del cliente, la ubicacion y nombre del plan de trabajo.
    //9.Enviar Email a Contabilidad informando sobre la activacion del nuevo cliente, la documentacion legal correspondiente y la informacion sobre facturacion.
    
    
    }
  