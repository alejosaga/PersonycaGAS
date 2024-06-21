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
  
    const otroSheet = SpreadsheetApp.openById(cotApproveId).getSheetByName('Aprobaciones');
    const emailSheet = SpreadsheetApp.openById(maestroCotId).getSheetByName('Datos'); // Asegúrate de que el nombre de la hoja es correcto
  
    if (!otroSheet || !emailSheet) {
      Logger.log('No se encontró una de las hojas especificadas.');
      return;
    }
  
    // Obtener todos los datos del sheet de formularios y del sheet de emails
    const dataRange = otroSheet.getDataRange();
    const data = dataRange.getValues();
    const emailDataRange = emailSheet.getDataRange();
    const emailData = emailDataRange.getValues();
  
    cotizaciones.forEach(cotizacion => {
      const cotizacionValue = cotizacion['Cotización'];
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
  
        // Buscar el email correspondiente usando el mismo índice
        const clienteNit = cotizacionValue.split(' NIT ')[1];
        let email = null;
        emailData.forEach(emailRow => {
          if (emailRow.includes(clienteNit)) {
            email = emailRow[1]; // Asumiendo que el email está en la columna B (índice 1)
          }
        });
  
        if (email) {
          // Obtener el archivo de Google Drive usando la URL
          const fileId = fileUrl.split('/d/')[1].split('/')[0]; // Extraer el ID del archivo de la URL
          const file = DriveApp.getFileById(fileId);
  
          // Enviar el correo electrónico con el archivo adjunto
          MailApp.sendEmail({
            to: email,
            subject: 'Archivo adjunto para cotización',
            body: `Por favor, revise el archivo adjunto para la cotización ${cotizacionValue}.`,
            attachments: [file.getAs(MimeType.PDF)]
          });
          Logger.log(`Correo enviado a ${email} con archivo adjunto para ${cotizacionValue}`);
        } else {
          Logger.log(`No se encontró email para el cliente NIT ${clienteNit}`);
        }
      } else {
        Logger.log(`No se encontró resultado para ${cotizacionValue}`);
      }
    });
  }