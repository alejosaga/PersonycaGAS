function traerDatos() {
  try {
    // Recupera datos Hoja de Aprobaciones
    let SSmaestroApprove = SpreadsheetApp.openById(cotApproveId);
    let sheetApprove = SSmaestroApprove.getSheetByName("Aprobaciones");
    let approvelastRow = sheetApprove.getLastRow();
    let approval = sheetApprove.getRange(approvelastRow, 7).getValue();
    let nit = sheetApprove.getRange(approvelastRow, 3).getValue();
    let razonSocial = sheetApprove.getRange(approvelastRow, 4).getValue();
    let numCot = sheetApprove.getRange(approvelastRow, 5).getValue();
    let sheetCot = sheetApprove.getRange(approvelastRow, 2).getValue();
    let valor = sheetApprove.getRange(approvelastRow, 6).getValue();
    
    Logger.log('SheetCot: ' + sheetCot);
    Logger.log('NumCot: ' + numCot);
  
    let parts = numCot.split("-");
    let numClien;
    if (parts[3] == "7") {
      numClien = parts[5];
    } else {
      numClien = parts[3];
    }
  
    let ssId;
    let servicio;
    switch(sheetCot) {
      case "Consultoria SG-Seguridad y salud en el trabajo":
        ssId = sgsstServiceId;
        servicio = searchValues(maestroCotId, numCot, "Datos", "cotizacion", "Servicios de interes");
        break;
      case "Aplicacion Bateria riesgo psicosocial":
        ssId = batPsiServiceId;
        servicio = searchValues(maestroCotId, numCot, "Datos", "cotizacion", "Servicios de interes");
        break;
      case "Sistemas de Gestion de Calidad -ISO":
        ssId = isoSisGesCotId;
        servicio = searchValues(maestroCotId, numCot, "Datos", "cotizacion", "De acuerdo a sus necesidades seleccione el sistema de gestión sobre el cual requiere consultoría");
        break;
      default:
        throw new Error("Sheet name not recognized: " + sheetCot);
    }
    
    Logger.log('ssId: ' + ssId);
  
    // Recupera datos Archivo Clientes
       
    let slideName = numCot;
    
    let slideId = searchValues(ssId, slideName, sheetCot, "slideName", "slideId");
    if (!slideId) {
      throw new Error("Slide ID not found for slide name: " + slideName);
    }
  
    let pdfFolderId = searchValues(ssId, slideName, sheetCot, "slideName", "IdcarpetaPdf");
    if (!pdfFolderId) {
      throw new Error("PDF Folder ID not found for slide name: " + slideName);
    }
  
    let contactName = searchValues(maestroCotId, numCot, "Datos", "cotizacion", "Nombres y apellidos de la persona contacto");
    if (!contactName) {
      throw new Error("Contact name not found for NIT: " + nit);
    }
  
    let carpetaPDF = DriveApp.getFolderById(pdfFolderId);
    let fechaHoy = new Date();
    let diasHabiles = 3; // Número de días hábiles a sumar
    
    let contador = 0;
    let fechaSumada = new Date(fechaHoy); // Use a new Date object to avoid modifying fechaHoy
    
    while (contador < diasHabiles) {
      fechaSumada.setDate(fechaSumada.getDate() + 1);
      
      // Si la fecha sumada no es un sábado (6) ni un domingo (0), se considera como día hábil
      if (fechaSumada.getDay() !== 6 && fechaSumada.getDay() !== 0) {
        contador++;
      }
    }
    
    let dia = fechaSumada.getDate();
    let mes = fechaSumada.getMonth() + 1; // El mes comienza desde 0, por lo que se suma 1
    let anio = fechaSumada.getFullYear();
    
    let diaFormateado = dia.toString().padStart(2, '0'); // Agrega un cero a la izquierda si el día es menor que 10
    let mesFormateado = mes.toString().padStart(2, '0'); // Agrega un cero a la izquierda si el mes es menor que 10
    let fechaFormateada = anio + "-" + mesFormateado + "-" + diaFormateado;
    
    Logger.log(fechaFormateada); // Imprime la fecha formateada en dd/mm/aaaa
    
    // Editar prefilled form -datos para diligenciar contrato-
    let cot = slideName.replace(/ /g, '+');
    let companyName = razonSocial.replace(/ /g, '+');
  
    let prefilledForm = "https://docs.google.com/forms/d/e/" + clientAcceptForm + "/viewform?usp=pp_url&entry.1204079246=" + nit + "&entry.219423794=" + cot + "&entry.1368409936=" + companyName + "&entry.1361289729=" + valor + "&entry.237567784=Si&entry.483030226=" + fechaFormateada;
  
      
    let pdfBlob = DriveApp.getFileById(slideId).getAs(MimeType.PDF);
    let pdfFile = carpetaPDF.createFile(pdfBlob);
    let fileId = pdfFile.getId();
    let pdffileUrl = pdfFile.getUrl();
    sheetApprove.getRange(approvelastRow, 8).setValue(pdffileUrl);
  
    let str = contactName;
    let firstWord = firstWordToTitleCase(str);
  
    let file = DriveApp.getFileById(fileId);
    let attach = file.getAs(MimeType.PDF); // Obtiene el archivo como un tipo de archivo específico (PDF en este caso);
  
    if (approval == "Si") {
      let subject = "Cotizacion " + servicio + " Personyca";
      let body = '<p>Ref: ' + slideName + '</p><p>Buen día Sr(a) <strong>' + firstWord + '</strong>,</p><p> Deseamos éxitos en sus actividades.</p><p>Gracias por elegirnos en conocer nuestros servicios para cubrir las necesidades de su empresa <strong>' + razonSocial + '</strong> en temas de <strong>' + servicio + '<strong>.</p> <p>como aliado estratégico de su organización en el servicio de consultoría en el diseño e implementación de sistemas de gestion. Contamos con licencia jurídica 2243 emitida por la secretaria de salud de Bogotá.</p><p>A continuación, encontrará nuestra oferta de servicios, esperamos que la misma satisfaga su propósito, no obstante estaremos atentos de suplir sus necesidades.</p><p>De nuevo le agradecemos su confianza en el equipo de PERSONYCA S.A.S.</p><p>A con Agradecemos su atención y solcitamos su colaboracion llenando el siguiente formulario indicandonos si acepta la cotizacion y proporcionando los datos necesarios para la elaboracion del contrato. ' + prefilledForm + ' </p><p>Cordialmente,</p></p><p>Equipo Personyca,</p>';
      MailApp.sendEmail({
        to: clientEmail1,
        cc: `${personycaEmail1},${personycaEmail2},${personycaEmail3}`,
        subject: subject,
        attachments: [attach],
        htmlBody: body
      });
    }
  } catch (e) {
    Logger.log('Error: ' + e);
    throw e;
  }
}
