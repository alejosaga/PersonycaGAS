function traerDatos() {
   //recupera datos Hoja de Aprobaciones
   var SSmaestroApprove = SpreadsheetApp.openById(cotApproveId);
   var sheetApprove = SSmaestroApprove.getSheetByName("Aprobaciones");
   var approvelastRow = sheetApprove.getLastRow();
   var approval = sheetApprove.getRange(approvelastRow,7).getValue();
   var nit = sheetApprove.getRange(approvelastRow,3).getValue();
   var razonSocial = sheetApprove.getRange(approvelastRow,4).getValue();
   var numCot = sheetApprove.getRange(approvelastRow,5).getValue();
   var sheetCot = sheetApprove.getRange(approvelastRow,2).getValue();
   var valor = sheetApprove.getRange(approvelastRow,6).getValue();
  
 
   var parts = numCot.split("-");
   
   if(parts[3]== "7"){
     numClien = parts[5]
   }
   else{
     numClien = parts[3]
   }
 
  
  switch(sheetCot) {
    case "Consultoria SG-Seguridad y salud en el trabajo":
      ssId = batPsiServiceId;
      
    case "Aplicacion Bateria riesgo psicosocial":
      ssId = sgsstServiceId;
  }
 
   
  //Recupera datos Archivo Clientes
  var servicio =  searchValues(maestroCotId,nit,"Datos","Nit","Servicios de interes");   
  var slideName = numCot;
  
  var slideId = searchValues(ssId,slideName,sheetCot,"slideName","slideId");
  //var slide = DriveApp.getFileById(slideId).getUrl();


  var pdfFolderId = searchValues(ssId,slideName,sheetCot,"slideName","IdcarpetaPdf");
 
  var contactName = searchValues(maestroCotId,nit,"Datos","Nit","Nombres y apellidos de la persona contacto");
  var carpetaPDF = DriveApp.getFolderById(pdfFolderId);
  var fechaHoy = new Date();
  var diasHabiles = 3; // Número de días hábiles a sumar
  
  var contador = 0;
  var fechaSumada = fechaHoy;
  
  while (contador < diasHabiles) {
    fechaSumada.setDate(fechaSumada.getDate() + 1);
    
    // Si la fecha sumada no es un sábado (6) ni un domingo (0), se considera como día hábil
    if (fechaSumada.getDay() !== 6 && fechaSumada.getDay() !== 0) {
      contador++;
    }
  }
  
  var dia = fechaSumada.getDate();
  var mes = fechaSumada.getMonth() + 1; // El mes comienza desde 0, por lo que se suma 1
  var anio = fechaSumada.getFullYear();
  
  var diaFormateado = dia.toString().padStart(2, '0'); // Agrega un cero a la izquierda si el día es menor que 10
  var mesFormateado = mes.toString().padStart(2, '0'); // Agrega un cero a la izquierda si el mes es menor que 10
  var fechaFormateada = anio + "-" + mesFormateado + "-" + diaFormateado;
  
  Logger.log(fechaFormateada); // Imprime la fecha formateada en dd/mm/aaaa
  
  
 //editar prefilled form -datos para diligenciar contrato-
 var cot = slideName.replace(/ /g, '+');
 var companyName = razonSocial.replace(/ /g, '+');

 

  var prefilledForm = "https://docs.google.com/forms/d/e/1FAIpQLSebS5fS5oB9tXnns6OdM6r8RnkC9PzYiU-CoptypDl-mhldPg/viewform?usp=pp_url&entry.1204079246="+nit+"&entry.219423794="+cot+"&entry.1368409936="+companyName+"&entry.1361289729="+valor+"&entry.237567784=Si&entry.483030226="+fechaFormateada;

 

  sheetApprove.getRange(approvelastRow,8).setValue(prefilledForm);


  var pdfBlob = DriveApp.getFileById(slideId).getAs(MimeType.PDF);
  var pdfFile = carpetaPDF.createFile(pdfBlob);
  var fileId= pdfFile.getId();
  
  var str = contactName;
  var firstWord = firstWordToTitleCase(str);
  
  var file = DriveApp.getFileById(fileId);
  var attach = file.getAs(MimeType.PDF); // Obtiene el archivo como un tipo de archivo específico (PDF en este caso);
  
  if(approval=="Si"){
  var subject = "Cotizacion "+servicio+" Personyca";
  var body = '<p>Ref: '+slideName+'</p><p>Buen día Sr(a) <strong>'+firstWord+'</strong>,</p><p> Deseamos éxitos en sus actividades.</p><p>Gracias por elegirnos en conocer nuestros servicios para cubrir las necesidades de su empresa <strong>'+razonSocial+'</strong> en temas de <strong>'+servicio+'<strong>.</p> <p>como aliado estratégico de su organización en el servicio de consultoría en el diseño e implementación de sistemas de gestion. Contamos con licencia jurídica 2243 emitida por la secretaria de salud de Bogotá.</p><p>A continuación, encontrará nuestra oferta de servicios, esperamos que la misma satisfaga su propósito, no obstante estaremos atentos de suplir sus necesidades.</p><p>De nuevo le agradecemos su confianza en el equipo de PERSONYCA S.A.S.</p><p>A con Agradecemos su atención y solcitamos su colaboracion llenando el siguiente formulario indicandonos si acepta la cotizacion y proporcionando los datos necesarios para la elaboracion del contrato. '+prefilledForm+' </p><p>Cordialmente,</p></p><p>Equipo Personyca,</p>' ; 
 MailApp.sendEmail({
  to: clientEmail1,clientEmail2,
  cc: personycaEmail,
  subject: subject,
  attachments: [attach],
  htmlBody: body
});

}

}




