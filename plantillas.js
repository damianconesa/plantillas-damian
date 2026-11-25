/*
 * ============================================================
 *  PLANTILLAS — edita este archivo para cambiar las plantillas
 * ============================================================
 *
 * Cada plantilla tiene:
 *   id          identificador interno (sin espacios)
 *   nombre      lo que ve quien te escribe
 *   descripcion una línea explicando cuándo usarla
 *   asunto      asunto del correo; {{idCampo}} se sustituye por lo que escriba la persona
 *   campos      lista de campos del formulario
 *
 * Tipos de campo: "texto", "parrafo" (varias líneas), "fecha", "opciones" (desplegable).
 * Pon obligatorio: true en los que no se puedan dejar vacíos.
 *
 * Tras guardar y publicar, el cambio llega a todo el mundo sin volver a desplegar el complemento.
 */
var PLANTILLAS = [
  {
    id: "incidencia",
    nombre: "Incidencia / soporte",
    descripcion: "Algo no funciona o da error en una aplicación o sistema.",
    asunto: "[Incidencia][{{urgencia}}] {{sistema}} – {{resumen}}",
    campos: [
      { id: "resumen", etiqueta: "Resumen en una frase", tipo: "texto", obligatorio: true,
        ayuda: "Ej.: No se genera la factura del contrato 1234" },
      { id: "sistema", etiqueta: "Aplicación o sistema", tipo: "opciones", obligatorio: true,
        opciones: ["Business Central", "Power Apps", "Power Automate", "Power BI", "Outlook / Teams", "Otro"] },
      { id: "descripcion", etiqueta: "¿Qué ocurre?", tipo: "parrafo", obligatorio: true,
        ayuda: "Qué intentabas hacer, qué esperabas y qué ha pasado. Copia el mensaje de error si lo hay." },
      { id: "pasos", etiqueta: "Pasos para reproducirlo", tipo: "parrafo", obligatorio: false },
      { id: "desde", etiqueta: "¿Desde cuándo pasa?", tipo: "fecha", obligatorio: false },
      { id: "urgencia", etiqueta: "Urgencia", tipo: "opciones", obligatorio: true,
        opciones: ["Baja", "Media", "Alta", "Bloqueante"] }
    ]
  },
  {
    id: "peticion",
    nombre: "Petición nueva / mejora",
    descripcion: "Necesitas algo nuevo o un cambio en algo que ya existe.",
    asunto: "[Petición][{{prioridad}}] {{titulo}}",
    campos: [
      { id: "titulo", etiqueta: "Título de la petición", tipo: "texto", obligatorio: true,
        ayuda: "Ej.: Informe mensual de suministros por edificio" },
      { id: "area", etiqueta: "Área o departamento", tipo: "texto", obligatorio: true },
      { id: "necesidad", etiqueta: "¿Qué necesitas?", tipo: "parrafo", obligatorio: true },
      { id: "objetivo", etiqueta: "¿Para qué? (beneficio esperado)", tipo: "parrafo", obligatorio: true,
        ayuda: "Qué ahorra, qué mejora o qué problema resuelve." },
      { id: "fecha", etiqueta: "Fecha deseada", tipo: "fecha", obligatorio: false },
      { id: "prioridad", etiqueta: "Prioridad", tipo: "opciones", obligatorio: true,
        opciones: ["Baja", "Media", "Alta"] }
    ]
  }
];
