/**
 * Función para enviar un email masivo siguiente criterios y filtros.
 * En el primer bloque a continuación, ingrese los valores solicitados, para ver un videotutorial en la charla, ingrese a:
 * https://drive.google.com/file/d/1-sBHawjUXM8TyP403cBJZg98w8rc0_r0/view?usp=sharing 
 */

var depuracion = 1;                         // 0 para enviar, 1 para enviar una muestra a correo personal

var nombreHoja = "myGoogleSheet"; // Entre las comillas, digite el nombre de la hoja donde están los datos, ejemplo "Hoja 1"
const numFilas = 17731;                          // Ingrese el número de la última fila en la que hay un contacto
const numColumnas = 31;                        // Ingrese el número de columnas en la tabla de contactos
const columnaEmails = "E-mail";               // Escriba el nombre de la columna donde se encuentran los correos de los contactos
const columnaNombre = "Primer Nombre";                      // Escriba el nombre de la columna donde se encuentra el nombre del contacto
const codigoHTML = "myHTML";        // Ingrese el nombre del archivo html que contiene el contenido del correo

  /** Hoja donde se encuentra la información de los estudiantes y rango de los datos*/ 
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja); 
const dataRange = sheet.getRange(2, 1, numFilas, numColumnas);
const rangeTitles = sheet.getRange(1, 1, 1, numColumnas);

const colEmail = getColumnNr(columnaEmails);
const colNombre = getColumnNr(columnaNombre);
const colEstado = getColumnNr(columnaEstado);

var correoPrueba = "heyderpaez@gmail.com";  // Escriba un correo de su uso o acceso, para revisar los correos de prueba
var remitente = "Correo Masivo";             // Coloque el nombre que desee aparezca como remitente, ejemplo, "Charlas Empresariales"
var responderA = "dir_investigacion@pca.edu.co"   // Coloque la dirección donde desea se redirijan las respuestas de los receptores
const enableRespuesta = true;              // Si desea enviarle como correo que no se puede responder (true) o si acepta respuesta (false)

var asunto = "Correo de prueba";            // Ingrese el asunto para enviar en el correo

/** Plantilla del Mensaje - Plantilla HTML del correo masivo*/
var bodyMail = HtmlService.createHtmlOutputFromFile(codigoHTML); 


var contactos = dataRange.getValues();

function sendMailCharla() {
  var enviados = 0;

  for (var i in contactos) {
    var contacto = contactos[i];

    var nombre = contacto[columnaNombre-1];
    var sendTo = contacto[columnaEmails-1];

    if (depuracion==1){
      var sendTo = correoPrueba;
      var preAsunto = "CORREO DE PRUEBA: ";
    }
    else{
      var preAsunto = "";
    }

    if(sendTo!="")  {
      var message = bodyMail.getContent();

      message = message.replace("%nombre", nombre);
   
      var subject = preAsunto + nombre + ", " + asunto;
      
      try {
        GmailApp.sendEmail(sendTo, subject, message, {htmlBody : message, name : remitente, replyTo:responderA, noReply:enableRespuesta });
      } catch (e) {
        // Logs an ERROR message.
        console.error('Se produjo el siguiente error al enviar el correo a '+ sendTo + ': ' + e);
        {break;}
      }    
      enviados=enviados+1;
      if(depuracion==1){break;}
    }     
  } 
  Logger.log("Enviados: " + enviados);
  Logger.log("Último envío: " + sendTo);
}

function getColumnNr(name) {
  var values = rangeTitles.getValues();
  
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == name) {
        return parseInt(col);
      }
    }
  }
  throw 'failed to get column by name';
}

