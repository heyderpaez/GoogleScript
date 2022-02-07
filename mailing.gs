/**
 * Función para enviar un email masivo siguiente criterios y filtros.
 * En el primer bloque a continuación, ingrese los valores solicitados, para ver un videotutorial en la charla, ingrese a:
 * https://drive.google.com/file/d/1-sBHawjUXM8TyP403cBJZg98w8rc0_r0/view?usp=sharing 
 */

var depuracion = 1;                         // 0 para enviar, 1 para enviar una muestra a correo personal

var nombreHoja = "Respuestas de formulario 1"; // Entre las comillas, digite el nombre de la hoja donde están los datos, ejemplo "Hoja 1"
var numFilas = 23;                          // Ingrese el número de la última fila en la que hay un contacto
var numColumnas = 8;                        // Ingrese el número de columnas en la tabla de contactos
var columnaEmails = 2;                      // Escriba el número de la columna donde se encuentran los correos de los contactos
var columnaNombre = 3;                      // Escriba el número de la columna donde se encuentra el nombre del contacto
var codigoHTML = "mailGoogleScript";        // Ingrese el nombre del archivo html que contiene el contenido del correo

var correoPrueba = "heyderpaez@gmail.com";  // Escriba un correo de su uso o acceso, para revisar los correos de prueba
var remitente = "Correo Masivo";             // Coloque el nombre que desee aparezca como remitente, ejemplo, "Charlas Empresariales"
var responderA = "dir_investigacion@pca.edu.co"   // Coloque la dirección donde desea se redirijan las respuestas de los receptores
var enableRespuesta = false;              // Si desea enviarle como correo que no se puede responder (false) o si acepta respuesta (true)

var asunto = "Correo de prueba";            // Ingrese el asunto para enviar en el correo

/** Plantilla del Mensaje - Plantilla HTML del correo masivo*/
var bodyMail = HtmlService.createHtmlOutputFromFile(codigoHTML); 

/** Hoja donde se encuentra la información de los estudiantes y rango de los datos*/ 
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja); 
var dataRange = sheet.getRange(2, 1, numFilas, numColumnas);

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
        
      GmailApp.sendEmail(sendTo, subject, message, {htmlBody : message, name : remitente, replyTo:responderA, noReply:enableRespuesta });
      enviados=enviados+1;
      if(depuracion==1){break;}
    }   
  } //Comentarear para producción
  Logger.log("Enviados: " + enviados);
}

