function prefillForm() {
    const sheetId = '1f5LwW6Ko0o4mUVrhgO5Fa6AiiRTdmENyhx4Oj8r3hK0';
    const formId = '1q0gnfJRANe7t6JEtqlXpvCqm02ooUHoICuRnJ0MeP6c'; // Reemplaza con el ID de tu formulario de Google
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Aprobaciones'); // Asegúrate de que el nombre de la hoja sea correcto
    const form = FormApp.openById(formId);
    
    // Borra las opciones actuales
    const items = form.getItems(FormApp.ItemType.CHECKBOX);
    const checkboxItem = items[0].asCheckboxItem();
    checkboxItem.setChoiceValues([]);
  
    // Obtener las cotizaciones del último mes
    const oneMonthAgo = new Date();
    oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
    const data = sheet.getDataRange().getValues();
  
    let options = [];
    data.forEach((row, index) => {
      if (index === 0) return; // Saltar la fila de encabezado
      const dateSent = new Date(row[0]);
      if (dateSent >= oneMonthAgo) {
        options.push(`Fecha: ${row[0]}, Cliente: ${row[3]}, Cotización: ${row[4]}`);
      }
    });
  
    // Añadir nuevas opciones
    checkboxItem.setChoiceValues(options);
  }
  
  function sendWeeklyReminder() {
    const emailRecipient = 'autopersonyca@gmail.com'; // Reemplaza con tu email
    const formUrl = 'https://docs.google.com/forms/d/1q0gnfJRANe7t6JEtqlXpvCqm02ooUHoICuRnJ0MeP6c/edit'; // Reemplaza con la URL de tu formulario
  
    prefillForm();
  
    MailApp.sendEmail({
      to: emailRecipient,
      subject: 'Recordatorio semanal de cotizaciones',
      body: `Lista de cotizaciones enviadas durante el último mes. Puedes seleccionar las cotizaciones para reenviar el recordatorio en el siguiente enlace: ${formUrl}`,
    });
  }
  
  function setupTrigger() {
    ScriptApp.newTrigger('sendWeeklyReminder')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .create();
  }
  