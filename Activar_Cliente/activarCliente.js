async function activarCliente() {
    let ssmasterContractApprove = SpreadsheetApp.openById(contractApproveId);
    let sheetAprContr = ssmasterContractApprove.getSheetByName("Datos");
    let lastRowContApr = sheetAprContr.getLastRow();
    let contrato = sheetAprContr.getRange(lastRowContApr, 3).getValue();
    let nombreRazonSocial = sheetAprContr.getRange(lastRowContApr, 6).getValue();
    
    
    let numClien = searchValues(maestroCotId,contrato,"Datos","cotizacion","Codigo Cliente");
    
    let clientEmail_1 = searchValues(maestroCotId,contrato,"Datos","cotizacion","Dirección de correo electrónico");
    let clientEmail_2 = searchValues(maestroCotId,contrato,"Datos","cotizacion","Segundo correo electronico (opcional)");


    
    
    let service = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","Servicios de interes")
  
    //2.Agregar fecha de inicio del contrato en la columna 7
      let fecIni = searchValues(contractmaestroId,numClien,"Datos","Codigo Cliente","Fecha de inicio");
      sheetAprContr.getRange(lastRowContApr,7).setValue(fecIni);
  
    //3.Buscar contrato en la carpeta del Cliente en "Prospectos".
  
      let folderId = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","clientFolderId");
      
      sheetAprContr.getRange(lastRowContApr,10).setValue(service);
  
    //4.Convertir Contrato a PDF y dejarlo en la carpeta del cliente prospecto.
  
      let contratoUrl = searchValues(contractmaestroId,numClien,"Datos","Codigo Cliente","urlContrato");
      let contractFolderId = searchValues(contractmaestroId,numClien,"Datos","Codigo Cliente","Carpeta Contratos");
      let contratoId = getIdFromUrl(contratoUrl);
      let archivoPDFId = convertirDocAPDF(contratoId, contractFolderId);  
  
     //5.Trasladar carpeta del cliente desde "Prospectos" a "Activos".
  
      let newfolderId = trasladarCarpeta(folderId, folderClienteActivoId);
      sheetAprContr.getRange(lastRowContApr,8).setValue(newfolderId);
  
  
    //6.Enviar contrato firmado por Personyca al Cliente mediante Email.

      let email = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","Dirección de correo electrónico");
      let subject = " Contrato adjunto - Confirmar recepción y firma: " + contrato;
      let archivoAdjunto = DriveApp.getFileById(archivoPDFId);
      //sendEmail(subject,toClient,email,archivoAdjunto)
      let nombreCliente = searchValues(maestroCotId,numClien,"Datos","Codigo Cliente","Nombres y apellidos de la persona contacto")
      let dataToSend = '<p>Estimado(a),</p> <strong>'+ nombreCliente +'</strong>,</p><p>Adjunto encontrará el contrato: ' +contrato+ ' que hemos preparado para usted. Por favor, revise el contrato detenidamente y, si está de acuerdo con los términos, le agradeceríamos que lo firmara y nos lo devolviera por correo electrónico.</p><p>Si tiene alguna pregunta o inquietud con respecto al contrato, no dude en comunicarse con nosotros. Estaremos encantados de brindarle cualquier aclaración adicional que necesite.</p><p>Una vez recibido el contrato firmado, procederemos de inmediato con los trámites necesarios para avanzar en el proceso. Esperamos con entusiasmo comenzar esta relación de trabajo y colaborar juntos.</p><p>Adjunto encontrará el contrato firmado en formato PDF. Si prefiere recibirlo en otro formato o necesita asistencia adicional, por favor háganoslo saber.</p><p>Agradecemos su atención y cooperación en este asunto. Quedamos a su disposición para cualquier consulta adicional.</p><p>Saludos cordiales,</p><p>Nancy Camacho<br>Gerente<br>Personyca SAS<br>personycasas@gmail.com<br>3165549102</p>'; 

      let body = dataToSend;
      MailApp.sendEmail({
          to: clientEmail_1,
          cc: clientEmail_2,
          subject: subject,
          htmlBody: body,
          attachments: [archivoAdjunto.getAs(MimeType.PDF)]
        }); 

    //1.Crear espacio carpeta y lista en clickUp
    main(nombreRazonSocial, sheetAprContr, lastRowContApr, contrato);
    
        
    // Obtener el ID del espacio, ya sea existente o creado
    
    //let space1= getSpaceId(nombreRazonSocial) 
    //console(space1)

  

     
    //createFolder(espacio, contrato);



    // Manejar el resultado
   
    
  
      
  
    //7.Guardar el plan de trabajo en carpeta del Cliente en caso de que se cree mediante Sheets.
    //8.Enviar Email a gerente de Personyca y consultores vinculados al proyecto, informando sobre la activacion del cliente, la ubicacion y nombre del plan de trabajo.
    //9.Enviar Email a Contabilidad informando sobre la activacion del nuevo cliente, la documentacion legal correspondiente y la informacion sobre facturacion.
    
    
    }
  