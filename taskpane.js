/* Panel de plantillas: muestra las dos plantillas, genera el formulario y lo vuelca al correo. */
(function () {
  "use strict";

  var CLAVE_AVISO = "plantillasDamian";
  var CLAVE_SESION = "plantillaDamianAplicada";

  var plantillaActual = null;
  var enOutlook = false;

  function $(id) { return document.getElementById(id); }

  function mostrarVista(nombre) {
    ["lista", "form", "ok"].forEach(function (v) {
      $("vista-" + v).hidden = v !== nombre;
    });
  }

  function escaparHtml(texto) {
    return String(texto)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function formatearFecha(valor) {
    var p = valor.split("-");
    return p.length === 3 ? p[2] + "/" + p[1] + "/" + p[0] : valor;
  }

  /* ---------- Lista de plantillas ---------- */

  function pintarLista() {
    var cont = $("lista-plantillas");
    cont.innerHTML = "";
    PLANTILLAS.forEach(function (p) {
      var b = document.createElement("button");
      b.type = "button";
      b.className = "tarjeta";
      b.innerHTML = "<span class='tarjeta-nombre'>" + escaparHtml(p.nombre) + "</span>" +
                    "<span class='tarjeta-desc'>" + escaparHtml(p.descripcion) + "</span>";
      b.addEventListener("click", function () { abrirPlantilla(p); });
      cont.appendChild(b);
    });
  }

  /* ---------- Formulario ---------- */

  function crearCampo(c) {
    var envoltorio = document.createElement("div");
    envoltorio.className = "campo";

    var label = document.createElement("label");
    label.htmlFor = "f_" + c.id;
    label.textContent = c.etiqueta + (c.obligatorio ? " *" : "");
    envoltorio.appendChild(label);

    var input;
    if (c.tipo === "parrafo") {
      input = document.createElement("textarea");
      input.rows = 4;
    } else if (c.tipo === "opciones") {
      input = document.createElement("select");
      var vacia = document.createElement("option");
      vacia.value = "";
      vacia.textContent = "Selecciona…";
      input.appendChild(vacia);
      (c.opciones || []).forEach(function (o) {
        var opt = document.createElement("option");
        opt.value = o;
        opt.textContent = o;
        input.appendChild(opt);
      });
    } else {
      input = document.createElement("input");
      input.type = c.tipo === "fecha" ? "date" : "text";
    }
    input.id = "f_" + c.id;
    input.name = c.id;
    if (c.obligatorio) input.required = true;
    envoltorio.appendChild(input);

    if (c.ayuda) {
      var ayuda = document.createElement("small");
      ayuda.textContent = c.ayuda;
      envoltorio.appendChild(ayuda);
    }
    return envoltorio;
  }

  function abrirPlantilla(p) {
    plantillaActual = p;
    $("form-titulo").textContent = p.nombre;
    var campos = $("campos");
    campos.innerHTML = "";
    p.campos.forEach(function (c) { campos.appendChild(crearCampo(c)); });
    $("aviso-error").hidden = true;
    mostrarVista("form");
    var primero = campos.querySelector("input, textarea, select");
    if (primero) primero.focus();
  }

  function leerValores() {
    var valores = {};
    var faltan = [];
    plantillaActual.campos.forEach(function (c) {
      var el = $("f_" + c.id);
      var v = (el.value || "").trim();
      el.classList.remove("invalido");
      if (c.obligatorio && !v) {
        faltan.push(c.etiqueta);
        el.classList.add("invalido");
      }
      valores[c.id] = c.tipo === "fecha" && v ? formatearFecha(v) : v;
    });
    return { valores: valores, faltan: faltan };
  }

  /* ---------- Generación del correo ---------- */

  function construirAsunto(p, valores) {
    return p.asunto.replace(/\{\{(\w+)\}\}/g, function (_, id) {
      return valores[id] || "";
    }).replace(/\[\]/g, "").replace(/\s+/g, " ").replace(/\s+–\s*$/, "").trim();
  }

  function construirCuerpo(p, valores) {
    var filas = p.campos.map(function (c) {
      var v = valores[c.id];
      var contenido = v ? escaparHtml(v).replace(/\r?\n/g, "<br>") : "<span style='color:#8a8886'>—</span>";
      return "<tr>" +
        "<td style='padding:6px 12px 6px 0;vertical-align:top;color:#605e5c;white-space:nowrap;border-bottom:1px solid #edebe9;'>" +
          escaparHtml(c.etiqueta) + "</td>" +
        "<td style='padding:6px 0;vertical-align:top;color:#201f1e;border-bottom:1px solid #edebe9;'>" + contenido + "</td>" +
        "</tr>";
    }).join("");

    return "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:11pt;'>" +
      "<p>Hola Damián,</p>" +
      "<p style='margin:0 0 6px 0;'><strong>" + escaparHtml(p.nombre) + "</strong></p>" +
      "<table cellpadding='0' cellspacing='0' style='border-collapse:collapse;font-size:10.5pt;margin-bottom:12px;'>" +
      filas + "</table>" +
      "<p>Gracias.</p></div><br>";
  }

  function aplicarAlCorreo(p, valores) {
    var item = Office.context.mailbox.item;
    var asunto = construirAsunto(p, valores);
    var cuerpo = construirCuerpo(p, valores);

    return new Promise(function (resolve, reject) {
      item.subject.getAsync(function (resAsunto) {
        var asuntoActual = resAsunto.status === Office.AsyncResultStatus.Succeeded ? (resAsunto.value || "").trim() : "";
        // En respuestas (RE:/RV:) se mantiene el asunto del hilo.
        var esRespuesta = /^(re|rv|fw|fwd|reenv)\s*:/i.test(asuntoActual);
        var cambiarAsunto = !esRespuesta;

        // El contenido se añade al principio para no borrar la firma ni el hilo.
        item.body.prependAsync(cuerpo, { coercionType: Office.CoercionType.Html }, function (resCuerpo) {
          if (resCuerpo.status !== Office.AsyncResultStatus.Succeeded) {
            reject(resCuerpo.error);
            return;
          }
          var terminar = function () {
            marcarAplicada(item);
            resolve({ asuntoCambiado: cambiarAsunto });
          };
          if (cambiarAsunto) {
            item.subject.setAsync(asunto, function () { terminar(); });
          } else {
            terminar();
          }
        });
      });
    });
  }

  function marcarAplicada(item) {
    try {
      if (item.sessionData) item.sessionData.setAsync(CLAVE_SESION, "1", function () {});
    } catch (e) { /* sin soporte: no pasa nada */ }
    try {
      item.notificationMessages.removeAsync(CLAVE_AVISO, function () {});
    } catch (e) { /* sin aviso que quitar */ }
  }

  /* ---------- Eventos ---------- */

  function alEnviarFormulario(e) {
    e.preventDefault();
    var r = leerValores();
    var aviso = $("aviso-error");
    if (r.faltan.length) {
      aviso.textContent = "Falta rellenar: " + r.faltan.join(", ") + ".";
      aviso.hidden = false;
      return;
    }
    aviso.hidden = true;

    if (!enOutlook) {
      // Vista previa fuera de Outlook: muestra el resultado en consola.
      console.log("Asunto:", construirAsunto(plantillaActual, r.valores));
      console.log("Cuerpo:", construirCuerpo(plantillaActual, r.valores));
      $("ok-texto").textContent = "Vista previa: abre el panel dentro de Outlook para añadirlo al correo.";
      mostrarVista("ok");
      return;
    }

    var boton = $("btn-insertar");
    boton.disabled = true;
    aplicarAlCorreo(plantillaActual, r.valores).then(function (res) {
      $("ok-texto").textContent = res.asuntoCambiado
        ? "Se han rellenado el asunto y el cuerpo. Revisa el correo y envíalo cuando quieras."
        : "Se ha añadido el contenido al cuerpo (se mantiene el asunto de la respuesta). Revisa y envía.";
      mostrarVista("ok");
    }).catch(function (err) {
      aviso.textContent = "No se pudo añadir al correo" + (err && err.message ? ": " + err.message : ".");
      aviso.hidden = false;
    }).then(function () { boton.disabled = false; });
  }

  function iniciar(info) {
    enOutlook = !!(info && info.host === Office.HostType.Outlook);
    pintarLista();
    mostrarVista("lista");
    $("formulario").addEventListener("submit", alEnviarFormulario);
    $("btn-volver").addEventListener("click", function () { mostrarVista("lista"); });
    $("btn-otra").addEventListener("click", function () { mostrarVista("lista"); });
  }

  if (window.Office && Office.onReady) {
    Office.onReady(iniciar);
  } else {
    document.addEventListener("DOMContentLoaded", function () { iniciar(null); });
  }
})();
