function buscarTarifas(caracteristicas) {
  let data = sheetTarifas.getDataRange().getValues();
  let headers = data[0]; 
  let conditionIndex = headers.indexOf("VlrBuscar");
  let valueIndex = headers.indexOf("peso");

  let valoresEncontrados = []; // Nueva matriz para almacenar los valores encontrados

  let condicionesEncontradas = data.filter(function(row) {
    
    return caracteristicas.includes(row[conditionIndex]);
  });

  // Agregar valores encontrados a la matriz de valores encontrados
  condicionesEncontradas.forEach(function(row) {
    valoresEncontrados.push(row[valueIndex]);
  });
  
  // Devolver la matriz de valores encontrados después de la iteración del ciclo
  if (valoresEncontrados.length > 0) {
    return valoresEncontrados;
  }

  // Si no se encuentra ningún valor, devolver "null" después de la iteración del ciclo
  Logger.log("No se encontró ningún valor para las condiciones ingresadas");
  return null;
}


function addRowNumber(SServicio,sheetName,columnNumber) {
 
  let sheet = SServicio.getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues();
  let nextNumber = 1;
  for (let i = 1; i < data.length; i++) { // start at 1 to skip the header row
    sheet.getRange(i+1, columnNumber).setValue(nextNumber); // write the next number in column C of the current row
    nextNumber++;
  }
}
function crearCarpetaCot(servi) {
  
  // Obtén el valor de la columna deseada de la última fila
  let folderName = sheetDatos.getRange(lastRowDat, 6).getValue(); // en este caso es la columna 6
   // Valida si ya existe una carpeta con el mismo nombre
  let existeCarpeta = carpetaRaiz.getFoldersByName(folderName).hasNext();
  
  // Crea la carpeta con el nombre del cliente en la carpeta prospectos
  if (!existeCarpeta) {
    let folder = carpetaRaiz.createFolder(folderName);
    let newFolder = DriveApp.getFoldersByName(folderName).next();
    let clientFolderId = newFolder.getId();
    let clientFolderURL = newFolder.getUrl();
    sheetDatos.getRange(lastRowDat,lastColumnDat-1).setValue(clientFolderId);
    sheetDatos.getRange(lastRowDat,lastColumnDat).setValue(clientFolderURL);
    // Obtén las subcarpetas dentro del folder inicial
    let newFolder1 = newFolder.createFolder(servi);
    let subFolder1 = newFolder1.createFolder("Cotizaciones");
    subFolder1.createFolder("PDF")
    subFolder1.createFolder("DOC")
    
    
    let folderId = newFolder1.getId();
    let folderURL = subFolder1.getUrl();
    return {id: folderId, url: folderURL};
    
  }else{
    let folderExists = false;
  
  // Obtén las subcarpetas dentro del folder inicial
  let inicialFolder = carpetaRaiz.getFoldersByName(folderName).next();
  let inicialFolderId = inicialFolder.getId();
  let inicialFolderUrl = inicialFolder.getUrl();
  let subfolders = inicialFolder.getFolders();
    while (subfolders.hasNext()) {
      let subfolder = subfolders.next();
      if (subfolder.getName() == servi) {
        folderExists = true;
        let folderId = subfolder.getId();
        let folder = subfolder.getFoldersByName("Cotizaciones").next();
        let folderURL = folder.getUrl();

        sheetDatos.getRange(lastRowDat,lastColumnDat-1).setValue(inicialFolderId);
        sheetDatos.getRange(lastRowDat,lastColumnDat).setValue(inicialFolderUrl);
           
        return {id: folderId, url: folderURL};              
        
      }   
    }  
    if (!folderExists) {  
      let newFolder = inicialFolder.createFolder(servi);
      let subFolder1 = newFolder.createFolder("Cotizaciones");
      subFolder1.createFolder("PDF")
      subFolder1.createFolder("DOC")
      
      
      let folderId = newFolder.getId();
      let folderURL = subFolder1.getUrl();
      
    }
  }
  
  return {id: folderId, url: folderURL};
}

function obtenerActividad(ciiu){

  let ciiuEmp = ciiu
  let searchValue = ciiuEmp;
  let sheet = sheetCiiu;
  let data = sheet.getDataRange().getValues();
    
    let result = "";
    result += "<table border='1' cellspacing='0' cellpadding='5'>" +
              "<tr>" +
              "<th>Codigo CIIU</th>" +
              "<th>CODIGO2</th>" +
              "<th>CLASE DE RIESGO</th>" +
              "<th>DESCRIPCIÓN DE ACTIVIDAD ECONÓMICA FINAL</th>" +
              "</tr>";
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] == searchValue) {
        result += "<tr><td>" + searchValue + "</td>" +
                  "<td>" + data[i][2] + "</td>" +
                  "<td>" + data[i][0] + "</td>" +
                  "<td>" + data[i][3] + "</td></tr>";
      }
    }
    result += "</table>";
    if (result.indexOf("<tr>") == -1) {
      Browser.msgBox("No results found for the search value: " + searchValue);
    
    }
    return result;
}


function preFilledForm(total,sheetCotizaciones,lastRowCot,column){
  //Form aprobacion pre-llenado
  let slideName = sheetCotizaciones.getRange(lastRowCot+1,column).getValue();
  console.log(slideName);
  let numCot = slideName.replace(/ /g, '+');
  let hojaCot = servicio.replace(/ /g, '+');
  let companyName = razonSocial.replace(/ /g, '+');
  let prefilledForm= "https://docs.google.com/forms/d/e/"+approveCotForm+"/viewform?usp=pp_url&entry.1149107413="+hojaCot+"&entry.1538120679="+nit+"&entry.127145366="+numCot+"&entry.1514100276="+total+"&entry.933752610="+companyName

  return prefilledForm;
}

function htmlData(libro,hoja,col1,col2) {
    let sheet = libro.getSheetByName(hoja)
    let lastRow = sheet.getLastRow();
    let lastColumn = sheet.getLastColumn();
    let dataRange = sheet.getRange(lastRow, col1, lastRow, col2);
    let data = dataRange.getValues();
    let headers = sheet.getRange(1, col1, 1, col2).getValues();
    let formResponse = "<table style='width:100%; font-size:14px; border: 1px solid black; border-collapse: collapse;'><tr style='background-color: lightgray;'><th style='border: 1px solid black; padding: 10px;'>Pregunta</th><th style='border: 1px solid black; padding: 10px;'>Respuesta</th></tr>";
    for (let i = 0; i < data[0].length; i++) {
      formResponse += "<tr style='border: 1px solid black;'><td style='border: 1px solid black; padding: 10px;'>" + headers[0][i] + "</td><td style='border: 1px solid black; padding: 10px;'>" + data[0][i] + "</td></tr>";
    }
    formResponse += "</table>";
    return formResponse
}

function formatoColombiano(valor) {
  return new Intl.NumberFormat("es-CO", {
    style: "currency",
    currency: "COP"
  }).format(valor);
}
let numeroALetras = (function() {
    // Código basado en el comentario de @sapienman
    // Código basado en https://gist.github.com/alfchee/e563340276f89b22042a
    function Unidades(num) {

        switch (num) {
            case 1:
                return 'UN';
            case 2:
                return 'DOS';
            case 3:
                return 'TRES';
            case 4:
                return 'CUATRO';
            case 5:
                return 'CINCO';
            case 6:
                return 'SEIS';
            case 7:
                return 'SIETE';
            case 8:
                return 'OCHO';
            case 9:
                return 'NUEVE';
        }

        return '';
    } //Unidades()

    function Decenas(num) {

        let decena = Math.floor(num / 10);
        let unidad = num - (decena * 10);

        switch (decena) {
            case 1:
                switch (unidad) {
                    case 0:
                        return 'DIEZ';
                    case 1:
                        return 'ONCE';
                    case 2:
                        return 'DOCE';
                    case 3:
                        return 'TRECE';
                    case 4:
                        return 'CATORCE';
                    case 5:
                        return 'QUINCE';
                    default:
                        return 'DIECI' + Unidades(unidad);
                }
            case 2:
                switch (unidad) {
                    case 0:
                        return 'VEINTE';
                    default:
                        return 'VEINTI' + Unidades(unidad);
                }
            case 3:
                return DecenasY('TREINTA', unidad);
            case 4:
                return DecenasY('CUARENTA', unidad);
            case 5:
                return DecenasY('CINCUENTA', unidad);
            case 6:
                return DecenasY('SESENTA', unidad);
            case 7:
                return DecenasY('SETENTA', unidad);
            case 8:
                return DecenasY('OCHENTA', unidad);
            case 9:
                return DecenasY('NOVENTA', unidad);
            case 0:
                return Unidades(unidad);
        }
    } //Unidades()

    function DecenasY(strSin, numUnidades) {
        if (numUnidades > 0)
            return strSin + ' Y ' + Unidades(numUnidades)

        return strSin;
    } //DecenasY()

    function Centenas(num) {
        let centenas = Math.floor(num / 100);
        let decenas = num - (centenas * 100);

        switch (centenas) {
            case 1:
                if (decenas > 0)
                    return 'CIENTO ' + Decenas(decenas);
                return 'CIEN';
            case 2:
                return 'DOSCIENTOS ' + Decenas(decenas);
            case 3:
                return 'TRESCIENTOS ' + Decenas(decenas);
            case 4:
                return 'CUATROCIENTOS ' + Decenas(decenas);
            case 5:
                return 'QUINIENTOS ' + Decenas(decenas);
            case 6:
                return 'SEISCIENTOS ' + Decenas(decenas);
            case 7:
                return 'SETECIENTOS ' + Decenas(decenas);
            case 8:
                return 'OCHOCIENTOS ' + Decenas(decenas);
            case 9:
                return 'NOVECIENTOS ' + Decenas(decenas);
        }

        return Decenas(decenas);
    } //Centenas()

    function Seccion(num, divisor, strSingular, strPlural) {
        let cientos = Math.floor(num / divisor)
        let resto = num - (cientos * divisor)

        let letras = '';

        if (cientos > 0)
            if (cientos > 1)
                letras = Centenas(cientos) + ' ' + strPlural;
            else
                letras = strSingular;

        if (resto > 0)
            letras += '';

        return letras;
    } //Seccion()

    function Miles(num) {
        let divisor = 1000;
        let cientos = Math.floor(num / divisor)
        let resto = num - (cientos * divisor)

        let strMiles = Seccion(num, divisor, 'UN MIL', 'MIL');
        let strCentenas = Centenas(resto);

        if (strMiles == '')
            return strCentenas;

        return strMiles + ' ' + strCentenas;
    } //Miles()

    function Millones(num) {
        let divisor = 1000000;
        let cientos = Math.floor(num / divisor)
        let resto = num - (cientos * divisor)

        let strMillones = Seccion(num, divisor, 'UN MILLON ', 'MILLONES ');
        let strMiles = Miles(resto);

        if (strMillones == '')
            return strMiles;

        return strMillones + ' ' + strMiles;
    } //Millones()

    return function NumeroALetras(num, currency) {
        currency = currency || {};
        let data = {
            numero: num,
            enteros: Math.floor(num),
            centavos: (((Math.round(num * 100)) - (Math.floor(num) * 100))),
            letrasCentavos: '',
            letrasMonedaPlural: currency.plural || 'PESOS CHILENOS', //'PESOS', 'Dólares', 'Bolíletes', 'etcs'
            letrasMonedaSingular: currency.singular || 'PESO CHILENO', //'PESO', 'Dólar', 'Bolilet', 'etc'
            letrasMonedaCentavoPlural: currency.centPlural || 'CHIQUI PESOS CHILENOS',
            letrasMonedaCentavoSingular: currency.centSingular || 'CHIQUI PESO CHILENO'
        };

        if (data.centavos > 0) {
            data.letrasCentavos = 'CON ' + (function() {
                if (data.centavos == 1)
                    return Millones(data.centavos) + ' ' + data.letrasMonedaCentavoSingular;
                else
                    return Millones(data.centavos) + ' ' + data.letrasMonedaCentavoPlural;
            })();
        };

        if (data.enteros == 0)
            return 'CERO ' + data.letrasMonedaPlural + ' ' + data.letrasCentavos;
        if (data.enteros == 1)
            return Millones(data.enteros) + ' ' + data.letrasMonedaSingular + ' ' + data.letrasCentavos;
        else
            return Millones(data.enteros) + ' ' + data.letrasMonedaPlural + ' ' + data.letrasCentavos;
    };

})();
function getFolderIds(sgSstFolderId) {
  
  let sgSstFolder = DriveApp.getFolderById(sgSstFolderId);
  let cotizacionesFolder = sgSstFolder.getFoldersByName("Cotizaciones").next();
  let docFolder = cotizacionesFolder.getFoldersByName("DOC").next();
  let pdfFolder = cotizacionesFolder.getFoldersByName("PDF").next();
  return {doc: docFolder.getId(), pdf: pdfFolder.getId()};
  
}
function sendEmail(numCot,link,dataToSend) {

  let firstName = "Nancy";
  let subject = "Revisar: " + numCot;
  let body = '<p>Hola <strong>'+ firstName +'</strong>, tenemos una nueva cotizacion por revisar para la empresa <strong>'+razonSocial+'</strong> para el servicio de '+servicio+'.</p> <p>Adjunto se encuentra el archivo PDF y un link donde podras encontrar el detalle de la cotizacion y si se requiere hacer los cambios que se consideren pertinentes.'+link+'<p>Tambien podras revisar los valores en el archivo maestro en la hoja de cotizaciones correspondiente</p>'+linkMaestro+'<p>Las siguientes son las respuestas al formulario de diagnostico: </p>'+dataToSend;
  MailApp.sendEmail({
      to: personycaEmail1,
      cc: personycaEmail2,
      subject: subject,
      htmlBody: body
    }); 
}


function searchValues(ssId, vlrBuscado, sheetName, colBuscada, colRespuesta) {
  let cacheKey = ssId + '_' + sheetName;
  let cachedData = {}; // Assuming you have some caching mechanism in place

  // Verificar si los datos ya están en caché
  if (!cachedData[cacheKey]) {
    let spreadSheet = SpreadsheetApp.openById(ssId);
    let sheet = spreadSheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    let data = sheet.getDataRange().getValues();
    cachedData[cacheKey] = data; // Guardar los datos en caché
  }
  
  let data = cachedData[cacheKey];
  let headers = data[0];
  let conditionIndex = headers.indexOf(colBuscada);
  let valueIndex = headers.indexOf(colRespuesta);

  let condition = vlrBuscado;

  // Recorrer las filas en orden ascendente
  for (let rowIndex = data.length - 1; rowIndex >= 0; rowIndex--) {
    let row = data[rowIndex];
    if (row[conditionIndex] == condition) {
      return row[valueIndex];
    }
  }

  Logger.log("No se encontró ningún valor para la condición ingresada");
  return null;
}

function firstWordToTitleCase(str) {
    let firstWord = str.split(" ")[0];
    return firstWord.charAt(0).toUpperCase() + firstWord.substring(1).toLowerCase();
    
  }
  function convertirFecha(fecha) {
    let fechaOriginal = fecha;
    let fecha1 = new Date(fechaOriginal);
    let options = { year: 'numeric', month: 'long', day: 'numeric' };
    let fechaConvertida = fecha1.toLocaleDateString('es-ES', options);
    fechaConvertida = fechaConvertida.charAt(0).toUpperCase() + fechaConvertida.slice(1);
    fechaConvertida = fechaConvertida.replace('De', 'de');
    return fechaConvertida;
    
  }
  function trasladarCarpeta(idCarpetaTrasladar, idCarpetaDestino) {
    let carpetaTrasladar = DriveApp.getFolderById(idCarpetaTrasladar);
    let carpetaDestino = DriveApp.getFolderById(idCarpetaDestino);
  
    let nuevaCarpeta = carpetaDestino.createFolder(carpetaTrasladar.getName());
    copiarContenido(carpetaTrasladar, nuevaCarpeta);
  
    //carpetaTrasladar.setTrashed(true);
    let newFolderId = nuevaCarpeta.getId();
    return newFolderId
  }
  
  function copiarContenido(carpetaOrigen, carpetaDestino) {
    let archivos = carpetaOrigen.getFiles();
    while (archivos.hasNext()) {
      let archivo = archivos.next();
      archivo.makeCopy(archivo.getName(), carpetaDestino);
    }
    
    let subCarpetas = carpetaOrigen.getFolders();
    while (subCarpetas.hasNext()) {
      let subCarpetaOrigen = subCarpetas.next();
      let nuevaSubCarpetaDestino = carpetaDestino.createFolder(subCarpetaOrigen.getName());
      copiarContenido(subCarpetaOrigen, nuevaSubCarpetaDestino);
    }
  }
  function convertirDocAPDF(idArchivo, idCarpetaDestino) {
    let archivo = DriveApp.getFileById(idArchivo);
    let carpetaDestino = DriveApp.getFolderById(idCarpetaDestino);
  
    let blobPDF = archivo.getAs('application/pdf');
    let nombrePDF = archivo.getName() + ".pdf";
    let archivoPDF = carpetaDestino.createFile(blobPDF).setName(nombrePDF);
  
    return archivoPDF.getId();
  }
  function getIdFromUrl(url) {
    let id = "";
    let match = url.match(/[-\w]{25,}/);
    if (match) {
      id = match[0];
    }
    return id;
  }

  function findSpaceByName(data, spaceName) {
    try {
      if (!data || !data.spaces || !Array.isArray(data.spaces)) {
        throw new Error('Invalid data format');
      }
  
      let found = data.spaces.some(space => space.name === spaceName);
      return found;
    } catch (error) {
      //console.error('Error finding space by name:', error);
      return false;
    }
  }
