function prefillForm() {
    let cotSheetId = cotApproveId;
    let maestroSheetId = maestroCotId;
    let formId = reminderForm;

    try {
        Logger.log(`ID de la hoja de cotizaciones: ${cotSheetId}`);
        Logger.log(`ID de la hoja de maestro: ${maestroSheetId}`);
        Logger.log(`ID del formulario: ${formId}`);

        let cotSheet = SpreadsheetApp.openById(cotSheetId).getSheetByName('Aprobaciones');
        let maestroSheet = SpreadsheetApp.openById(maestroSheetId).getSheetByName('Datos');

        if (!cotSheet) {
            Logger.log('No se encontró la hoja de aprobaciones. Asegúrate de que el ID de la hoja es correcto.');
            return;
        }

        if (!maestroSheet) {
            Logger.log('No se encontró la hoja de datos maestros. Asegúrate de que el ID de la hoja es correcto.');
            return;
        }

        let cotDataRange = cotSheet.getDataRange();
        let cotData = cotDataRange.getValues();

        let maestroDataRange = maestroSheet.getDataRange();
        let maestroData = maestroDataRange.getValues();

        if (cotData.length === 0) {
            Logger.log('El rango de datos de cotizaciones está vacío.');
            return;
        }

        if (maestroData.length === 0) {
            Logger.log('El rango de datos maestros está vacío.');
            return;
        }

        let form = FormApp.openById(formId);
        if (!form) {
            Logger.log('No se encontró el formulario. Asegúrate de que el ID del formulario es correcto.');
            return;
        }

        // Elimina todas las preguntas existentes
        let items = form.getItems();
        items.forEach(item => form.deleteItem(item));

        // Crear una nueva pregunta de tipo checkbox
        let checkboxItem = form.addCheckboxItem();
        checkboxItem.setTitle('Estas cotizaciones fueron enviadas el último mes, a cuales le envías un recordatorio');

        // Obtener las cotizaciones del último mes
        let oneMonthAgo = new Date();
        oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);

        let options = [];
        let cotizacionesInfo = [];

        cotData.forEach((row, index) => {
            if (index === 0) return; // Saltar la fila de encabezado
            let dateSent = new Date(row[0]);
            if (dateSent >= oneMonthAgo) {
                let option = `Fecha: ${row[0]}, Cliente: ${row[3]}, Cotización: ${row[4]}`;
                options.push(option);

                // Buscar datos de contacto del cliente en el archivo maestro usando NIT
                let clienteNit = row[2].toString().trim();
                let clienteInfo = maestroData.find(maestroRow => maestroRow[4].toString().trim() == clienteNit); // Asumiendo que la columna E en maestro tiene el NIT del cliente

                if (clienteInfo) {
                    let cotizacionInfo = {
                        fecha: row[0],
                        numero: row[4],
                        valor: row[5],
                        clienteNombre: clienteInfo[6], // Columna F en maestro
                        clienteCargo: clienteInfo[8], // Columna H en maestro
                        clienteTelefono: clienteInfo[9], // Columna I en maestro
                        clienteEmail: clienteInfo[1], // Columna B en maestro
                        clienteEmail2: clienteInfo[2] // Columna C en maestro
                    };
                    cotizacionesInfo.push(cotizacionInfo);
                } else {
                    Logger.log(`No se encontró información de contacto para el cliente con NIT: ${clienteNit}`);
                }
            }
        });

        if (options.length === 0) {
            Logger.log('No se encontraron cotizaciones del último mes.');
            return;
        }

        try {
            // Añadir nuevas opciones
            checkboxItem.setChoiceValues(options);
        } catch (e) {
            Logger.log(`Error al añadir opciones: ${e.message}`);
        }

        // Obtener la URL pública del formulario
        let formUrl = form.getPublishedUrl();

        return { formUrl, cotizacionesInfo };
    } catch (e) {
        Logger.log(`Error en la función prefillForm: ${e.message}`);
    }
}

function sendWeeklyReminder() {
    let result = prefillForm();
    if (!result) {
        Logger.log('No se pudo obtener la URL del formulario o no hay cotizaciones para enviar.');
        return;
    }

    let { formUrl, cotizacionesInfo } = result;

    if (!formUrl || cotizacionesInfo.length === 0) {
        Logger.log('No se pudo obtener la URL del formulario o no hay cotizaciones para enviar.');
        return;
    }

    // Crear el cuerpo del correo electrónico
    let emailBody = 'Lista de cotizaciones enviadas durante el último mes:\n\n';
    cotizacionesInfo.forEach(info => {
        emailBody += `Fecha: ${info.fecha}\nNúmero de Cotización: ${info.numero}\nValor: ${info.valor}\n`;
        emailBody += `Cliente: ${info.clienteNombre}\nCargo: ${info.clienteCargo}\nTeléfono: ${info.clienteTelefono}\nEmail: ${info.clienteEmail}\n\n`;
    });
    emailBody += `Puedes seleccionar las cotizaciones para reenviar el recordatorio en el siguiente enlace: ${formUrl}`;

    MailApp.sendEmail({
        to: personycaEmail1,
        cc: `${personycaEmail2},${personycaEmail3}`,
        subject: 'Recordatorio semanal de cotizaciones',
        body: emailBody,
    });
}

function setupTrigger() {
    ScriptApp.newTrigger('sendWeeklyReminder')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(9)
        .create();
}
