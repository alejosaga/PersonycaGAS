function generarCotizacionISO() {
    const BD_servicio = sgsstServiceId;
    const SServicio = SpreadsheetApp.openById(BD_servicio);
    const sheetCotizaciones = SServicio.getSheetByName(servicio);
    const lastRowCot = sheetCotizaciones.getLastRow();
    const lastColumnCot = sheetCotizaciones.getLastColumn()

   /* To continue....
  
    // Tarifas según el tipo de servicio
        let NumProcesos = 0;     
        let MaestroDocu =0; 
        let dirTec = 0;
        let dirTecExp = 0;
        let PrevAudi = 0;
        let capacitaciones = 0;
        let Consultor = 0;
        let costoAdmin = 0;
        let costoOperativo = 0;
        let rentPersonyca = 0;
  
    switch(consultoria){
        case "ISO - 9001 (Sistema de Gestión por procesos para la satisfacción del cliente)":
            NumProcesos = 0;   
            MaestroDocu =0; 
            dirTec = 0;
            dirTecExp = 0;
            PrevAudi = 0;
            capacitaciones = 0;
            Consultor = 0;
            costoAdmin = 0;
            costoOperativo = 0;
            rentPersonyca = 0;
        break;

        case "ISO - 13485 (Requisitos para propositos regulatorios dispositivos médicos)":
            NumProcesos = 0;   
            MaestroDocu =0; 
            dirTec = 0;
            dirTecExp = 0;
            PrevAudi = 0;
            capacitaciones = 0;
            Consultor = 0;
            costoAdmin = 0;
            costoOperativo = 0;
            rentPersonyca = 0;
        break;

        case "ISO - 45001 (Seguridad y salud en el trabajo)":
            NumProcesos = 0;   
            MaestroDocu =0; 
            dirTec = 0;
            dirTecExp = 0;
            PrevAudi = 0;
            capacitaciones = 0;
            Consultor = 0;
            costoAdmin = 0;
            costoOperativo = 0;
            rentPersonyca = 0;
        break;

        case "ISO - 14001 (Sistemas de gestión ambiental)":
            NumProcesos = 0;   
            MaestroDocu =0; 
            dirTec = 0;
            dirTecExp = 0;
            PrevAudi = 0;
            capacitaciones = 0;
            Consultor = 0;
            costoAdmin = 0;
            costoOperativo = 0;
            rentPersonyca = 0;
        break;

        case "Sistemas integrados de Gestión (9001, 45001 y 45001)":
            NumProcesos = 0;   
            MaestroDocu =0; 
            dirTec = 0;
            dirTecExp = 0;
            PrevAudi = 0;
            capacitaciones = 0;
            Consultor = 0;
            costoAdmin = 0;
            costoOperativo = 0;
            rentPersonyca = 0;
        break;

        case "Sistemas integrados de Gestión (9001 y 13485)":
            NumProcesos = 0;   
            MaestroDocu =0; 
            dirTec = 0;
            dirTecExp = 0;
            PrevAudi = 0;
            capacitaciones = 0;
            Consultor = 0;
            costoAdmin = 0;
            costoOperativo = 0;
            rentPersonyca = 0;
        break;
      default:
        throw new Error("Servicio no soportado");
    }
  
    const total = tarifaBasica + costosOperativos + marketing;
  
    // Insertar los datos de la cotización en la hoja de cotizaciones
    sheetCotizaciones.appendRow([
      nit,
      razonSocial,
      new Date(),
      tarifaBasica,
      costosOperativos,
      marketing,
      total
    ]);
  
    // Enviar el correo con la cotización
    const emailTemplate = HtmlService.createTemplateFromFile('plantilla_email');
    emailTemplate.razonSocial = razonSocial;
    emailTemplate.total = total;
  
    const emailBody = emailTemplate.evaluate().getContent();
  
    MailApp.sendEmail({
      to: 'client@example.com', // Reemplaza con el correo del cliente
      subject: `Cotización de ${servicio}`,
      htmlBody: emailBody
    });
  }
  
  // Ejemplo de uso:
  function cotizarISO9001() {
    generarCotizacionISO("Consultoria ISO 9001");
  }
  
  function cotizarISO14001() {
    generarCotizacionISO("Consultoria ISO 14001");
  }
  
  function cotizarISO45001() {
    generarCotizacionISO("Consultoria ISO 45001");*/
  }
  

