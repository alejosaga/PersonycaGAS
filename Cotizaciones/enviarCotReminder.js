function parseCotizacionString() {
  const SSreminders = SpreadsheetApp.openById(remidersCotId);
  const shetReminders = SSreminders.getSheetByName("Reminders");
  const lastRowRem = shetReminders.getLastRow();
  const lastColumnRem = shetReminders.getLastColumn();

  let data = shetReminders.getRange(lastRowRem, 2).getValue();  // Obtener el string de cotizaciones
  console.log(data);

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

  console.log(cotizacionesArray);
  return cotizacionesArray;
}

function buscarEnOtroSheet() {
  const cotizaciones = parseCotizacionString();

  const otroSheet = SpreadsheetApp.openById(batPsiServiceId).getSheetByName('Sheet1');

  if (!otroSheet) {
    Logger.log('No se encontró la hoja especificada.');
    return;
  }

  // Obtener todos los datos del sheet
  const dataRange = otroSheet.getDataRange();
  const data = dataRange.getValues();

  cotizaciones.forEach(cotizacion => {
    const cotizacionValue = cotizacion['Cotización'];

    // Buscar el valor de la cotización en la columna correspondiente
    let resultado;
    data.forEach(row => {
      if (row.includes(cotizacionValue)) {
        resultado = row;
      }
    });

    if (resultado) {
      Logger.log(`Resultado encontrado para ${cotizacionValue}: ${resultado}`);
      // Puedes hacer algo con el resultado aquí
    } else {
      Logger.log(`No se encontró el valor ${cotizacionValue} en el sheet.`);
    }
  });
}

// Ejemplo de uso
buscarEnOtroSheet();
