# Plantillas para Damián — complemento de Outlook

Cuando alguien de la empresa pone a Damián en **Para** o **CC**, Outlook le muestra un aviso encima del correo con el botón **Elegir plantilla**. Al pulsarlo se abre un panel con dos plantillas (*Incidencia / soporte* y *Petición nueva / mejora*), la persona rellena el formulario y el complemento escribe el asunto y el cuerpo del correo. También se puede abrir a mano desde el botón **Plantillas para Damián** de la cinta.

**Coste: 0 €.** Los archivos se alojan en GitHub Pages (gratis) y el despliegue se hace desde el Centro de administración de Microsoft 365, incluido en vuestra licencia.

---

## Qué hay en la carpeta

| Archivo | Para qué sirve |
|---|---|
| `manifest.xml` | Define el complemento. Es lo que se sube al Centro de administración. |
| `launchevent.js` | Detecta si Damián está entre los destinatarios y muestra el aviso. Aquí están sus direcciones. |
| `commands.html` | Carga `launchevent.js` en Outlook web, nuevo Outlook y Mac. |
| `plantillas.js` | **Las dos plantillas.** Es el único archivo que tendrás que tocar para cambiar campos o textos. |
| `taskpane.html / .js / .css` | El panel con las plantillas y el formulario. |
| `index.html` | Página de soporte que exige el manifiesto. |
| `assets/` | Iconos. |

---

## 1. Publicar los archivos gratis (GitHub Pages)

1. Crea en GitHub un repositorio **público** llamado `plantillas-damian` (GitHub Pages es gratis en repos públicos; en privados requiere plan de pago). No hay nada secreto en el código: solo tus direcciones de correo y los textos de las plantillas.
2. Sube **todo el contenido** de esta carpeta a la raíz del repositorio.
3. En el repositorio: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, rama `main`, carpeta `/ (root)` → **Save**.
4. En un par de minutos tendrás la web en `https://TU-damianconesa.github.io/plantillas-damian/`. Comprueba que abre `.../taskpane.html`.

> Alternativa también gratuita si preferís que quede en Azure: **Azure Static Web Apps, plan Free**.

## 2. Poner la URL real en el manifiesto

En `manifest.xml`, busca y reemplaza **todas** las apariciones de:

```
https://damianconesa.github.io/plantillas-damian
```

por tu URL real (sin barra final), y en `<AppDomain>` deja solo el dominio: `https://TU-damianconesa.github.io`. Vuelve a subir el `manifest.xml` al repo para tenerlo versionado.

## 3. Probarlo tú primero (sin afectar a nadie)

1. Abre Outlook en la web → **Nuevo correo** → menú **⋯ → Aplicaciones → Agregar aplicaciones** (o entra directamente en `https://aka.ms/olksideload`).
2. **Mis complementos → Complementos personalizados → Agregar un complemento personalizado → Agregar desde archivo** y elige `manifest.xml`.
3. Redacta un correo nuevo y ponte a ti mismo (`damianconesa@alvores.com`) como destinatario: debe aparecer el aviso **Elegir plantilla**.
4. Prueba las dos plantillas y revisa cómo queda el correo.

## 4. Desplegarlo a la empresa

Lo hace alguien con rol de **Administrador global** o **Administrador de Exchange**:

1. `https://admin.microsoft.com` → **Configuración → Aplicaciones integradas → Cargar aplicaciones personalizadas**.
2. Tipo: **Complemento de Office** → **Cargar archivo de manifiesto** → `manifest.xml`.
3. Asignar a **toda la organización** o a un grupo concreto (recomendado empezar por un grupo piloto).
4. Aceptar los permisos y terminar. Puede tardar **varias horas (hasta 24 h o algo más)** en aparecer en el Outlook de los usuarios.

Si lo probaste en el paso 3, quita antes tu versión personal para no tener el complemento duplicado.

## 5. Cambiar las plantillas

Edita `plantillas.js` en GitHub y guarda (commit). En unos minutos todos ven la versión nueva; no hay que volver a desplegar nada.

Solo hay que **volver a subir el manifiesto** si cambias la URL, el nombre del complemento o los eventos. En ese caso sube también el número de `<Version>`.

Para añadir otra dirección tuya que deba activar el aviso, añádela a `DESTINATARIOS_DAMIAN` en `launchevent.js`.

---

## Comportamiento y límites

- **Solo para personas de la organización** que tengan el complemento desplegado. A alguien de fuera no le aparece nada.
- Funciona en **Outlook web, nuevo Outlook para Windows, Outlook clásico para Windows (Microsoft 365 actualizado) y Mac**. En **móvil no se activa**: el aviso con botón no está disponible en Outlook para iOS/Android.
- No obliga a usar la plantilla: el aviso se puede cerrar y el correo se envía igual.
- El texto de la plantilla se **añade al principio** del cuerpo, así no se pierde la firma ni el hilo en las respuestas. En respuestas (`RE:`, `RV:`) se mantiene el asunto del hilo.
- Una vez aplicada una plantilla, el aviso deja de salir en ese correo.
