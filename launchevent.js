/*
 * Detección del destinatario.
 * Se ejecuta sin interfaz cada vez que cambian los destinatarios de un correo en redacción
 * (y al abrir un correo nuevo o una respuesta). Si Damián está en Para o CC, muestra un
 * aviso con el botón "Elegir plantilla", que abre el panel.
 *
 * Escrito en JavaScript sencillo (sin módulos ni sintaxis moderna) porque Outlook clásico
 * para Windows ejecuta este archivo en un entorno JS reducido.
 */

// Direcciones que activan el aviso (en minúsculas).
var DESTINATARIOS_DAMIAN = [
  "damianconesa@alvores.com",
  "damianconesa@azpool.es"
];

var CLAVE_AVISO = "plantillasDamian";
var CLAVE_SESION = "plantillaDamianAplicada";
var ID_BOTON = "btnPlantillasDamian"; // id del botón en manifest.xml

function esDamian(destinatario) {
  if (!destinatario || !destinatario.emailAddress) return false;
  return DESTINATARIOS_DAMIAN.indexOf(destinatario.emailAddress.toLowerCase()) !== -1;
}

function leerCampo(campo, callback) {
  if (!campo || !campo.getAsync) { callback([]); return; }
  campo.getAsync(function (res) {
    callback(res.status === Office.AsyncResultStatus.Succeeded && res.value ? res.value : []);
  });
}

function plantillaYaAplicada(item, callback) {
  try {
    if (!item.sessionData) { callback(false); return; }
    item.sessionData.getAsync(CLAVE_SESION, function (res) {
      callback(res.status === Office.AsyncResultStatus.Succeeded && res.value === "1");
    });
  } catch (e) {
    callback(false);
  }
}

function mostrarAviso(item, event) {
  item.notificationMessages.replaceAsync(CLAVE_AVISO, {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: "Escribes a Damián: usa una de sus plantillas para que tu solicitud llegue completa.",
    icon: "Icon.16x16",
    actions: [{
      actionText: "Elegir plantilla",
      actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
      commandId: ID_BOTON,
      contextData: "{\"origen\":\"aviso\"}"
    }]
  }, function () { event.completed(); });
}

function quitarAviso(item, event) {
  item.notificationMessages.removeAsync(CLAVE_AVISO, function () { event.completed(); });
}

function comprobarDestinatarios(event) {
  var item = Office.context.mailbox.item;
  leerCampo(item.to, function (para) {
    leerCampo(item.cc, function (cc) {
      var todos = para.concat(cc);
      var hayDamian = false;
      for (var i = 0; i < todos.length; i++) {
        if (esDamian(todos[i])) { hayDamian = true; break; }
      }
      if (!hayDamian) { quitarAviso(item, event); return; }
      plantillaYaAplicada(item, function (aplicada) {
        if (aplicada) { event.completed(); } else { mostrarAviso(item, event); }
      });
    });
  });
}

function onRecipientsChanged(event) {
  try { comprobarDestinatarios(event); } catch (e) { event.completed(); }
}

function onNewMessageCompose(event) {
  try { comprobarDestinatarios(event); } catch (e) { event.completed(); }
}

Office.actions.associate("onRecipientsChanged", onRecipientsChanged);
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
