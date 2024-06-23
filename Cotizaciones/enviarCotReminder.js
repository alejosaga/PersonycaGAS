function parseCotizacionString() {
    const SSreminders = SpreadsheetApp.openById(remidersCotId);
    const shetReminders = SSreminders.getSheetByName("Reminders");
    const lastRowRem = shetReminders.getLastRow();
  
    let data = shetReminders.getRange(lastRowRem, 2).getValue();  // Obtener el string de cotizaciones
  
    // Dividir el string completo en cotizaciones individuales
    const cotizaciones = data.split(', Fecha:').map((item, index) => index === 0 ? item : 'Fecha:' + item);
  
    let cotizacionesArray = [];
  
    // Recorrer cada cotización y dividir por ": " para obtener clave y valor
    cotizaciones.forEach(cotizacionString => {
      const parts = cotizacionString.split(',');
      const cotizacionData = {};
  
      parts.forEach(part => {
        const [key, value] = part.split(': ').map(item => item.trim());
        cotizacionData[key] = value;
      });
  
      cotizacionesArray.push(cotizacionData);
    });
  
    return cotizacionesArray;
  }
  
  function buscarEnOtroSheet() {
    const cotizaciones = parseCotizacionString();
  
    const otroSheet = SpreadsheetApp.openById(batPsiServiceId).getSheetByName('Aplicacion Bateria riesgo psico');
    const emailSheet = SpreadsheetApp.openById(maestroCotId).getSheetByName('Datos'); // Asegúrate de que el nombre de la hoja es correcto
    const aprobacionCotSheet = SpreadsheetApp.openById(cotApproveId).getSheetByName('Aprobaciones'); // Asegúrate de que el nombre de la hoja es correcto
  
    if (!otroSheet || !emailSheet || !aprobacionCotSheet) {
        Logger.log('No se encontró una de las hojas especificadas.');
        return;
      }
    
      // Obtener todos los datos del sheet de formularios y del sheet de emails
      const dataRange = otroSheet.getDataRange();
      const data = dataRange.getValues();
      const emailDataRange = emailSheet.getDataRange();
      const emailData = emailDataRange.getValues();
      const aprobacionCotDataRange = aprobacionCotSheet.getDataRange();
      const aprobacionCotData = aprobacionCotDataRange.getValues();
    
      cotizaciones.forEach(cotizacion => {
        let cotizacionValue = cotizacion['Cotización'];
        let resultado = null;
    
        // Buscar el valor de la cotización en la columna correspondiente
        data.forEach(row => {
          if (row.includes(cotizacionValue)) {
            resultado = row;
          }
        });
    
        if (resultado) {
          // Extraer la URL del archivo PDF de la última columna
          const fileUrl = resultado[resultado.length - 1]; // Última columna
          Logger.log(`URL del archivo PDF para ${cotizacionValue}: ${fileUrl}`);
    
          // Buscar el email, nombre de contacto y servicio de interés correspondiente usando la cotización
          let email = null;
          let nombreContacto = null;
          let servicioInteres = null;
          emailData.forEach(emailRow => {
            if (emailRow.includes(cotizacionValue)) {
              email = emailRow[1]; // Asumiendo que el email está en la columna B (índice 1)
              nombreContacto = emailRow[3]; // Asumiendo que el nombre de contacto está en la columna D (índice 3)
              servicioInteres = emailRow[5]; // Asumiendo que el servicio de interés está en la columna F (índice 5)
            }
          });
    
          // Buscar el enlace del formulario prellenado y la URL del archivo PDF en Aprobacion Cot-STG
          let formLink = null;
          aprobacionCotData.forEach(aprobacionRow => {
            if (aprobacionRow.includes(cotizacionValue)) {
              formLink = aprobacionRow[aprobacionRow.length - 2]; // Penúltima columna
              fileUrl = aprobacionRow[aprobacionRow.length - 1]; // Última columna
            }
          });
    
          if (email && formLink && fileUrl && fileUrl.includes('/d/')) {
            // Obtener el archivo de Google Drive usando la URL
            const fileId = fileUrl.split('/d/')[1].split('/')[0]; // Extraer el ID del archivo de la URL
            const file = DriveApp.getFileById(fileId);
    
            // Enviar el correo electrónico con el archivo adjunto
            MailApp.sendEmail({
              to: email,
              subject: `Seguimiento cotizacion ${servicioInteres} Personyca`,
              body: `Estimado(a) ${nombreContacto},\n\nEspero que se encuentre bien. Le escribimos para hacerle un seguimiento a la cotización ${cotizacionValue} de ${servicioInteres} que le enviamos recientemente.\n\nAdjunto encontrará el documento con todos los detalles que necesita. También hemos preparado un formulario que puede completar en el siguiente enlace: ${formLink}.\n\nEstamos a su disposición para cualquier duda o consulta que pueda tener. Nos encantaría poder ayudarle en todo lo que necesite y esperamos tener la oportunidad de trabajar con usted.\n\nAgradecemos su atención y quedamos atentos a su respuesta.\n\nSaludos cordiales,\nEl equipo de Personyca`,
              attachments: [file.getAs(MimeType.PDF)]
            });
            Logger.log(`Correo enviado a ${email} con archivo adjunto para ${cotizacionValue}`);
          } else {
            Logger.log(`No se encontró email, enlace del formulario, o URL válida para ${cotizacionValue}`);
          }
        } else {
          Logger.log(`No se encontró resultado para ${cotizacionValue}`);
        }
      });
    }