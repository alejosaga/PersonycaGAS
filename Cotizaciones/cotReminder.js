function prefillForm() {
    const sheetId = '1f5LwW6Ko0o4mUVrhgO5Fa6AiiRTdmENyhx4Oj8r3hK0';
    const formId = '1q0gnfJRANe7t6JEtqlXpvCqm02ooUHoICuRnJ0MeP6c';
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Aprobaciones');
    
   if (!sheet) {
    Logger.log('Hoja no encontrada. Asegúrate de que el nombre de la hoja es correcto.');
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  if (data.length === 0) {
    Logger.log('El rango de datos está vacío.');
    return;
  }
  
 
  
  const form = FormApp.openById(formId);
  
  // Elimina todas las preguntas existentes
  const items = form.getItems();
  items.forEach(item => form.deleteItem(item));
  
  // Crear una nueva pregunta de tipo checkbox
  const checkboxItem = form.addCheckboxItem();
  checkboxItem.setTitle('Estas cotizaciones fueron enviadas el último mes, a cuales le envías un recordatorio');
  
  // Obtener las cotizaciones del último mes
  const oneMonthAgo = new Date();
  oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
  
  let options = [];
  data.forEach((row, index) => {
    if (index === 0) return; // Saltar la fila de encabezado
    const dateSent = new Date(row[0]);
    Logger.log(`Fecha de envío: ${dateSent}`);
    if (dateSent >= oneMonthAgo) {
      const option = `Fecha: ${row[0]}, Cliente: ${row[3]}, Cotización: ${row[4]}`;
      options.push(option);
      Logger.log(`Opción agregada: ${option}`);
    }
  });
  
  Logger.log(`Opciones generadas: ${JSON.stringify(options)}`);
  
  if (options.length === 0) {
    Logger.log('No se encontraron cotizaciones del último mes.');
    return;
  }
  
  try {
    // Añadir nuevas opciones
    checkboxItem.setChoiceValues(options);
    Logger.log('Opciones añadidas al formulario.');
  } catch (e) {
    Logger.log(`Error al añadir opciones: ${e.message}`);
  }
}

function sendWeeklyReminder() {
  const emailRecipient = 'autopersonyca@gmail.com';
  const formUrl = 'https://docs.google.com/forms/d/1q0gnfJRANe7t6JEtqlXpvCqm02ooUHoICuRnJ0MeP6c/edit';
  
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
  