# Recrear el dashboard desde cero con GitHub Copilot

## 1. Objetivo y modo de uso

Esta guia es una especificacion autocontenida y una secuencia de prompts para construir una aplicacion equivalente, sin necesitar el codigo original. Incluye interfaz, datos, aprovisionamiento de SharePoint, autenticacion, escrituras, documentos, sincronizacion y pruebas.

**La unica entrada que debe cambiar quien reutiliza los prompts es `SITIO_SHAREPOINT`**, en el prompt inicial. Los nombres de listas, columnas, estados, puerto y comportamiento se mantienen. El tenant, el ID de cliente de Entra, el GUID de la lista relacionada y el ID del grupo de miembros son resultados del descubrimiento/aprovisionamiento: Copilot debe obtenerlos y configurar el proyecto, no pedir que se sustituyan manualmente en cada prompt.

Esto no significa que una URL otorgue permisos. Se necesita una cuenta autorizada para crear listas y columnas, registrar una SPA en Entra y consentir los permisos que permita la organizacion. Si Copilot no dispone de herramientas autenticadas para hacerlo, debe guiar la operacion en el portal y recoger sus resultados. El inicio de sesion, MFA, consentimiento y autorizaciones administrativas no se pueden omitir ni prometer como automaticos.

### Pasos para quien recibe esta guia

1. Crear una carpeta vacia y abrirla en VS Code con GitHub Copilot en modo agente.
2. Adjuntar **este archivo completo** a la conversacion.
3. Copiar el prompt 0, cambiar solo la URL y enviarlo.
4. Enviar los prompts 1 a 7 en orden, conservando el contexto de la conversacion y este archivo adjunto.
5. Completar las operaciones interactivas de Microsoft cuando corresponda.
6. Exigir el informe de verificacion final. Tener una interfaz que parece funcionar no demuestra que SharePoint este actualizado.

Los bloques de prompts remiten a las secciones de este mismo documento. En una conversacion nueva, volver a adjuntarlo completo; un prompt aislado sin esta especificacion no contiene todos los requisitos.

## 2. Alcance y arquitectura a reproducir

- Dashboard en espanol: **Gestion de Oportunidades Comerciales**.
- Aplicacion estatica en un unico `index.html`, con HTML, CSS y JavaScript nativo; sin framework, backend, base de datos propia ni compilacion.
- SharePoint Online es la fuente persistente de oportunidades, metadatos y archivos.
- Dos listas genericas: `Oportunidades` y `DocumentosOportunidades`. Los archivos son **adjuntos de elementos de lista**, no una biblioteca documental.
- REST de SharePoint mediante `fetch`. Microsoft Graph no se usa en el dashboard.
- OAuth delegado con MSAL Browser para localhost; sin secretos de cliente.
- Servidor de desarrollo `Start-Dashboard.ps1`, loopback, puerto 3000.
- Pruebas automatizadas sin dependencias externas usando `node:test`, al menos para miembros y sincronizacion.
- Sin datos ficticios al iniciar, sin guardado local de oportunidades/documentos y sin simulacion de exito ante errores.
- No crear Power Automate, webhooks, tareas programadas, busqueda global, edicion completa de oportunidades, roles de aprobacion ni validacion automatica de archivos.

### Dependencias y apariencia originales

| Componente | Referencia |
| --- | --- |
| Estilos | Tailwind mediante `https://cdn.tailwindcss.com` |
| Iconos | Lucide UMD `https://cdn.jsdelivr.net/npm/lucide@0.468.0/dist/umd/lucide.min.js` |
| Autenticacion | MSAL Browser `https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.27.0/lib/msal-browser.min.js`, carga al inicializar OAuth |
| Fuente | Segoe UI, system-ui, sans-serif |
| Colores | Azul `#0078D4`, azul oscuro `#106EBE`, navy `#002050`, verde `#107C10`, fondo azul claro `#EFF6FC` |
| Distribucion | Contenedor centrado de ancho maximo equivalente a `max-w-7xl`, fondo slate claro, paneles blancos, bordes y sombras suaves |

Conservar estas referencias para reproducir la demo. Si alguna deja de estar disponible, comunicarlo antes de sustituirla y verificar compatibilidad. Un despliegue de produccion requeriria dependencias empaquetadas y HTTPS administrado; no forma parte de esta replica local.

## 3. Descubrimiento y aprovisionamiento de SharePoint

### Reglas generales

1. Validar que la URL indicada corresponda a un sitio SharePoint Online accesible; normalizarla sin barra final. Derivar el origen desde la URL, sin asumir el dominio de otro entorno.
2. Descubrir el tenant real mediante metadatos de autenticacion de Microsoft/SharePoint, usando herramientas autorizadas. No deducirlo simplemente del nombre del host ni reutilizar un GUID de esta demo.
3. Leer el grupo de miembros asociado al sitio y su ID; no asumir que es el grupo 5.
4. Crear/verificar primero `Oportunidades` y despues `DocumentosOportunidades`, ambas listas genericas (plantilla 100).
5. Aprovisionar de manera idempotente: inspeccionar listas/campos existentes y reutilizar los compatibles. Si existe un campo incompatible, informar la diferencia y detener esa operacion; no borrar datos, cambiar nombres internos o recrear listas silenciosamente.
6. Para evitar nombres internos codificados, crear cada campo con su **nombre interno ASCII** y despues aplicar la etiqueta visible. `Title` ya existe: solo se renombra su titulo visible.
7. Conservar columnas nativas: ID, Created, Modified, Author, Editor y metadatos de version. No fabricarlas desde JavaScript.
8. Verificar la configuracion efectiva leyendo de nuevo los campos, vistas y propiedades. No basta con que una solicitud de creacion haya respondido correctamente.

### Lista `Oportunidades`

| Nombre interno | Etiqueta visible | Tipo y configuracion |
| --- | --- | --- |
| `Title` | Proyecto | Texto de una linea, obligatorio. La web limita a 120 caracteres. |
| `Cliente` | Cliente | Texto de una linea, obligatorio, maximo 80 caracteres. |
| `Presupuesto` | Presupuesto | Number, obligatorio, minimo 0, dos decimales. No Currency; la web presenta USD. |
| `Estado` | Estado | Choice, obligatorio, seleccion unica: `Prospecto`, `Negociación`, `Ganada`, `Perdida`. Sin opciones libres. |
| `Fecha` | Fecha | DateTime en modo solo fecha, obligatorio. |
| `AsignadoA` | Asignado a | User/Persona, seleccion unica, solo personas, **no obligatorio en SharePoint**. |

La tilde en `Negociación` es parte del valor persistido. No usar `Negociacion` en REST.

`AsignadoA` debe limitarse al grupo de miembros asociado: `SelectionGroup` igual al ID descubierto del grupo y `UserSelectionMode=PeopleOnly` (o propiedades equivalentes de la herramienta de aprovisionamiento). Agregarlo a la vista predeterminada.

La web exige responsable al crear, pero el campo de SharePoint es opcional para admitir registros historicos sin asignacion. No asignar una persona arbitrariamente al aprovisionar ni modificar registros existentes para completar ese campo.

### Lista `DocumentosOportunidades`

| Nombre interno | Etiqueta visible sugerida | Tipo y configuracion |
| --- | --- | --- |
| `Title` | Nombre del archivo | Texto de una linea; la web escribe el nombre completo del archivo, maximo 180 caracteres. |
| `Oportunidad` | Oportunidad | Lookup unico, obligatorio, dirigido al GUID real de `Oportunidades`, campo mostrado `Title`. Indexado y con `RelationshipDeleteBehavior=Restrict`. |
| `AsignadoA` | Asignado a | User/Persona, igual configuracion que en oportunidades, opcional, copia del responsable de la oportunidad. |
| `EstadoValidacion` | Estado de validacion | Choice unico: `Pendiente`, `Validado`, `Rechazado`; predeterminado `Pendiente`. Sin opciones libres. |
| `Observaciones` | Observaciones | Note, texto multilinea sin formato ni anexado de cambios. La web limita a 2000 caracteres. |
| `ValidadoPor` | Validado por | Texto de una linea para la cuenta/correo del revisor, no un campo Persona. Debe admitir null al volver a Pendiente. |
| `FechaValidacion` | Fecha de validacion | DateTime con fecha y hora; debe admitir null. Se escribe UTC. |

Configuracion adicional:

- Habilitar adjuntos (`EnableAttachments=true`) e historial de versiones (`EnableVersioning=true`).
- `AttachmentFiles` es la navegacion nativa de adjuntos, **no crear una columna personalizada con ese nombre**.
- Agregar `AsignadoA` a la vista predeterminada; mostrar tambien nombre, oportunidad y estado para facilitar la inspeccion.
- Comprobar que la restriccion de eliminacion funciona tambien al intentar borrar directamente desde SharePoint, no solo desde la web.
- Un elemento documental representa un archivo. La UI carga uno por operacion; una oportunidad admite varios elementos documentales.
- El historial de versiones de elementos ayuda a auditar metadatos. No prometer versionado binario de adjuntos equivalente al de una biblioteca documental.

Los tipos de texto de `ValidadoPor` y fecha/hora de `FechaValidacion` son el contrato necesario para las escrituras de la aplicacion. Esta guia se basa en el codigo y la documentacion del repositorio; no es una exportacion administrativa en vivo de todos los atributos de los campos del tenant original.

### Limitacion del selector nativo

El grupo de SharePoint puede contener un grupo Microsoft 365 en vez de personas directamente. El selector Persona nativo no siempre expande ese grupo anidado al buscar usuarios con `SelectionGroup` restringido. La replica debe asignar/reasignar desde la web y mostrar el resultado en SharePoint. **No** resolverlo ampliando permisos, agregando usuarios al grupo o permitiendo seleccionar todo el directorio.

## 4. Microsoft Entra ID y sesion

### Registro que debe crearse en el entorno destino

- Nombre sugerido: `demo-ia-with-copilot-dashboard`.
- Tipo de cuentas: solo este directorio, **single tenant**.
- Plataforma: **Single-page application / Aplicacion de pagina unica (SPA)**.
- Redirect URI principal exacta: `http://localhost:3000/`.
- Si se utilizara la ruta explicita, registrar tambien `http://localhost:3000/index.html`. La URI calculada por la web debe coincidir con una registrada.
- Permiso de API: **SharePoint > Delegated permissions > `AllSites.Write`**.
- Recurso SharePoint, no Microsoft Graph. Quitar el permiso predeterminado Graph `User.Read` si no se usa.
- No crear secreto/certificado para la SPA; no incluir credenciales en el repositorio.
- Mantener desactivados los flujos implicitos y la opcion de cliente publico.
- No conceder consentimiento global de administrador automaticamente. Si la politica lo requiere, detener el acceso hasta que un administrador autorizado lo apruebe.
- El usuario necesita permisos efectivos para leer y editar ambas listas y sus adjuntos. El permiso delegado no eleva sus permisos de SharePoint.

Los permisos para **aprovisionar** listas o registrar aplicaciones no son los mismos que los de ejecucion del dashboard. No agregar permisos administrativos a la SPA para facilitar la instalacion.

### Configuracion generada de la aplicacion

Centralizar en un objeto `CONFIG`:

```javascript
const CONFIG = {
  SITE_URL: /* URL suministrada en el prompt 0 */,
  TENANT_ID: /* tenant descubierto y verificado */,
  CLIENT_ID: /* Application (client) ID del registro creado o reutilizado */,
  LIST_TITLE: 'Oportunidades',
  DOCUMENTS_LIST_TITLE: 'DocumentosOportunidades'
};
```

Esto representa valores que Copilot debe completar, no un fragmento ejecutable para dejar con comentarios pendientes. Son identificadores publicos, no secretos. Derivar tambien los enlaces visibles a las listas desde la configuracion y las URL reales aprovisionadas, sin dejar enlaces al sitio original.

Comportamiento OAuth:

1. Crear `msal.PublicClientApplication`, inicializarlo, autoridad `https://login.microsoftonline.com/{TENANT_ID}`.
2. `redirectUri = location.origin + location.pathname`; no incluir query ni fragmento.
3. Scope: `{origenSharePoint}/AllSites.Write`.
4. Cache MSAL en `sessionStorage`.
5. Boton Conectar: validar GUID del cliente, `loginPopup` con `prompt: 'select_account'`, establecer cuenta activa.
6. Si el popup no puede abrirse (`popup_window_error` o `empty_window_error`), usar `loginRedirect`. No tratar una cancelacion voluntaria como bloqueo de popup.
7. Procesar `handleRedirectPromise` al iniciar y restaurar la cuenta activa para reconectar tras recarga.
8. Adquirir tokens con `acquireTokenSilent`. Ante interaccion requerida, marcar desconectado y solicitar Conectar de nuevo; no entrar en bucles de login.
9. Enviar `Authorization: Bearer ...` solo al REST del sitio configurado; no mandar tokens a enlaces externos ni a otro sitio recibido por paginacion.
10. Solo guardar en `localStorage` la preferencia `vista` y el ID publico `sharepointClientId`. No implementar guardado propio de tokens, documentos u oportunidades.

Estar autenticado en una pestana de SharePoint no autoriza la pagina localhost. Como alternativa de codigo, si la app se adapta a un entorno soportado del mismo origen, usar sesion del sitio y request digest (`/_api/contextinfo`, cache con expiracion y margen de 60 segundos). Subir el HTML a una biblioteca moderna no garantiza su ejecucion; no habilitar scripts personalizados ni debilitar la seguridad para conseguirlo.

## 5. Interfaz y reglas funcionales

### Cabecera y conexion

- Cabecera navy con distintivo O, nombre de la aplicacion y fecha actual larga en `es-ES` en escritorio.
- Campo visible para el ID de aplicacion Entra, prellenado con el registro del entorno destino.
- Botones **Conectar SharePoint**, **Actualizar**, enlace a lista de oportunidades y mensaje de conexion/error.
- Sin conexion: formulario y operaciones remotas deshabilitados, listado indica sin conexion, sin oportunidades precargadas.
- Bloquear controles mientras una operacion este en curso para evitar dobles envios.

### Indicadores

Calcular sobre **todas** las oportunidades cargadas, no sobre el filtro seleccionado:

| Indicador | Formula |
| --- | --- |
| Pipeline total | Suma de Presupuesto de todos los estados, incluidas Ganada y Perdida. |
| Ingresos ganados | Suma de Presupuesto con Estado = Ganada. |
| Abiertas | Cantidad con Estado Prospecto o Negociación. |
| Tasa de exito | Redondear `100 * ganadas / (ganadas + perdidas)`; 0% si no hay cerradas. |

Importes con `Intl.NumberFormat('es-ES', {style:'currency', currency:'USD', maximumFractionDigits:0})`. La presentacion sin decimales no debe redondear el dato guardado de dos decimales.

### Oportunidades

- Formulario: Cliente, Proyecto, Asignado a, Presupuesto (USD), Estado y Fecha.
- Cliente/proyecto recortados y no vacios, maximos 80/120. Fecha obligatoria, inicializada al dia de la implementacion mediante ISO (`new Date().toISOString().slice(0,10)`).
- Presupuesto obligatorio, finito, no negativo, representacion decimal con punto y hasta dos decimales; admitir cero y rechazar vacio, negativos y notacion exponencial.
- Estados exactos de la seccion 3; inicial Prospecto. Responsable obligatorio en la web.
- Guardar Proyecto en `Title`, fecha como `YYYY-MM-DDT00:00:00Z`, responsable como `AsignadoAId`.
- Tras alta confirmada: limpiar formulario, restablecer fecha, mostrar Todos y resaltar la nueva fila/tarjeta. Respetar movimiento reducido.
- Filtros Todos y cada estado con cantidades; botones Tabla/Tarjetas, recordar vista entre recargas.
- Orden visible por fecha descendente. Consulta inicial a SharePoint por Id descendente.
- Tabla: Cliente, Proyecto, Asignado a, Presupuesto, Estado, Fecha y acciones.
- Tarjetas: misma informacion, badge de estado, presupuesto destacado y las mismas acciones.
- En ambas vistas: selector de estado, selector de responsable, Eliminar con confirmacion, Documentos.
- No hay edicion de cliente, proyecto, presupuesto o fecha despues del alta.
- Al eliminar oportunidad, consultar antes si tiene documentos y bloquear si existen. La relacion Restrict debe proteger tambien en el servidor ante carreras u otros clientes.

### Responsables disponibles

- Leer exclusivamente `web/AssociatedMemberGroup/Users`, con paginacion.
- Personas directas: `PrincipalType === 1`.
- Para un grupo Microsoft 365, reconocer login `c:0o.c|federateddirectoryclaimprovider|{GUID}` y expandir mediante `SP.Directory.DirectorySession/Group('{GUID}')/Members`, tambien con paginacion.
- Esa API puede devolver propiedades `principalName`, `displayName`, `mail` en minuscula inicial aunque el `$select` use `PrincipalName,DisplayName,Mail`. Verificar la respuesta real antes de mapearla.
- Deduplicar por cuenta, eliminando el prefijo claims hasta el ultimo `|` y normalizando a minusculas; ordenar por nombre en espanol.
- Grupo no soportado, no enumerable, usuario sin cuenta o conjunto vacio: error explicito; no reemplazar por `siteusers` ni todo Entra.
- Volver a consultar pertenencia antes de crear o reasignar. Resolver el usuario con POST `web/ensureuser`, cuerpo `{logonName: cuenta}`, y usar el `Id` **del sitio SharePoint**.
- Los IDs de persona no son el correo, nombre visible ni Object ID de Entra.
- Una asignacion historica fuera del grupo se conserva y se muestra como ya no miembro; permitir elegir otro miembro. No ofrecer vaciar una asignacion desde el selector de la web.

### Panel documental

- Bajo oportunidades, con enlace a lista documental, selector de oportunidad y boton Actualizar documentos.
- Boton Documentos en fila/tarjeta selecciona la oportunidad y desplaza al panel.
- Resumen: cliente, proyecto, total de documentos y cantidad Pendiente.
- Carga de un archivo por operacion; varios documentos por oportunidad.
- Listado: nombre clicable para abrir **el formulario de revision**, responsable heredado, badge Pendiente/Validado/Rechazado, Descargar, Eliminar con confirmacion.
- Descargar con autenticacion a blob y enlace temporal; revocar URL posteriormente. No depender de abrir un enlace sin token.
- Revision: nombre, cuenta/fecha de revision, estado, observaciones y Guardar validacion.
- Rechazado exige observaciones no vacias tras `trim`; maximo 2000 caracteres para cualquier estado.
- Validado o Rechazado guardan cuenta activa y fecha actual UTC. En mismo origen, obtener cuenta desde `web/currentuser`. Pendiente limpia `ValidadoPor` y `FechaValidacion`, no el historial ni necesariamente las observaciones.
- Registro sin adjunto: mostrar Carga incompleta y deshabilitar descarga/revision; permitir eliminacion.
- Cambiar oportunidad limpia archivo seleccionado y revision anterior. Un fallo de lectura documental limpia resultados anteriores y no los muestra como actuales.
- No hay analisis de contenido, firmas, malware o cumplimiento. Cualquier usuario con edicion puede revisar; no hay separacion cargador/aprobador.
- `ValidadoPor` es un metadato editable, no prueba inmutable. Auditoria tambien con Editor, Modified y versiones nativas.

### Accesibilidad y responsive

- Labels asociados, estados con `role=status`/`aria-live`, alertas para errores, botones de icono con nombre accesible.
- Pantallas pequenas: formulario sobre listado, tarjetas adaptables y tablas con scroll horizontal **dentro** de su contenedor; no desbordar la pagina.
- Escritorio: formulario 1/3 y listado 2/3; panel documental con listado y revision a dos columnas en pantallas amplias.
- Escapar datos de SharePoint antes de insertarlos en HTML. No interpolar nombres de archivos/cuentas sin escape.

## 6. Contrato REST y actualizacion del sitio

Base: `{SITE_URL}/_api/`. Headers JSON: `Accept: application/json;odata=minimalmetadata`; cuerpos JSON con `Content-Type: application/json;odata=nometadata`. En localhost usar Bearer y `credentials:'omit'`; en mismo origen, cookies y digest para POST.

Validar cada URL, incluidas paginas siguientes: mismo origen y ruta bajo el `/_api/` del sitio configurado. No seguir una URL externa adjuntando un token. Escapar literales OData y codificar nombres de archivo, incluyendo apostrofes.

| Operacion | Solicitud relativa a `/_api/` |
| --- | --- |
| Leer oportunidades | GET `web/lists/getbytitle('Oportunidades')/items?$select=Id,Title,Cliente,Presupuesto,Estado,Fecha,AsignadoAId,AsignadoA/Title,AsignadoA/Name&$expand=AsignadoA&$orderby=Id desc&$top=500` |
| Crear oportunidad | POST a esa coleccion con `Title,Cliente,Presupuesto,Estado,Fecha,AsignadoAId`. |
| Cambiar estado/responsable | POST a `.../items(ID)`, `X-HTTP-Method: MERGE`, `IF-MATCH` con ETag leido y solo campos modificados. |
| Eliminar oportunidad | POST a `.../items(ID)`, `X-HTTP-Method: DELETE`, ETag exacto, tras comprobar documentos. |
| Leer documentos de oportunidad | GET `web/lists/getbytitle('DocumentosOportunidades')/items` filtrado por `$filter=OportunidadId eq ID`, orden Id desc, `$top=100`; seleccionar metadatos y expandir `AttachmentFiles,AsignadoA`. |
| Crear metadatos documentales | POST a coleccion documental con `Title`, `OportunidadId`, `AsignadoAId`, `EstadoValidacion:'Pendiente'`. |
| Adjuntar binario | POST `.../items(ID)/AttachmentFiles/add(FileName='NOMBRE_CODIFICADO')`, `application/octet-stream`, cuerpo ArrayBuffer del archivo. |
| Descargar | GET `.../items(ID)/AttachmentFiles/getByFileName('NOMBRE_CODIFICADO')/$value`, respuesta blob. |
| Guardar revision | MERGE del elemento documental con ETag, `EstadoValidacion,Observaciones,ValidadoPor,FechaValidacion`. |
| Eliminar documento | DELETE del elemento documental con ETag; elimina tambien su adjunto, nunca su oportunidad. |
| Sincronizar responsables | Leer todos los elementos documentales con `Id,OportunidadId,AsignadoAId` y ETag; MERGE solo de `AsignadoAId` cuando difiere. |

Para lectura documental seleccionar al menos `Id,Title,OportunidadId,AsignadoAId,AsignadoA/Title,EstadoValidacion,Observaciones,ValidadoPor,FechaValidacion,Created,AttachmentFiles`. Todos los lectores de colecciones deben seguir tanto `odata.nextLink` como `@odata.nextLink`. Obtener ETag de `odata.etag` o `@odata.etag` y conservarlo.

### Secuencia de escrituras y errores

- Actualizar la vista de datos como guardados solo despues de respuesta satisfactoria de SharePoint; no tratar un cambio optimista como confirmacion.
- Para MERGE/DELETE ordinarios exigir ETag exacto, nunca `IF-MATCH:*`. Despues de escribir, invalidar el ETag anterior y releer.
- HTTP 401: marcar desconexion. HTTP 403: indicar permisos de usuario/app. HTTP 404: lista/registro/archivo no encontrado. HTTP 412: otro usuario modifico, actualizar antes de reintentar. Mostrar tambien el detalle REST disponible.
- Distinguir **escritura fallida** de **escritura correcta con refresco fallido**. En el segundo caso decir que ya se guardo y no sugerir repetir el alta/carga.
- Conservar la proteccion de concurrencia si el refresco falla; no reutilizar un ETag invalidado.

### Carga y compensacion

1. Validar oportunidad positiva y entera, archivo presente y no vacio.
2. Limite exacto: `20 * 1024 * 1024` bytes; aceptar ese tamano y rechazar uno mayor.
3. Nombre hasta 180 caracteres; rechazar `" * : < > ? / \ |`, caracteres de control U+0000 a U+001F, punto/espacio inicial o final. Admitir apostrofes mediante escape OData correcto.
4. Leer los bytes y refrescar oportunidades antes de crear para obtener la asignacion actual; no usar la copia antigua del selector.
5. Crear elemento documental, luego enviar el binario. SharePoint no ofrece aqui una transaccion unica entre ambas solicitudes.
6. Si falla el binario, intentar borrar **solo el elemento provisional recien creado**. La implementacion original usa `IF-MATCH:*` exclusivamente en esta compensacion, no en ediciones ordinarias.
7. Si la limpieza falla o no se puede confirmar, informar el ID documental y pedir verificarlo antes de reintentar; no afirmar que no se guardo nada.
8. Si carga completa pero lectura posterior falla, indicar guardado correcto y pedir Actualizar documentos, no una nueva carga.

### Herencia y sincronizacion de responsables

La oportunidad es la fuente de verdad de `AsignadoA`. El documento contiene una copia de ese ID como metadato, no una modificacion de su archivo.

- Al conectar/restaurar sesion, al pulsar Actualizar y al pulsar Actualizar documentos: cargar miembros, cargar oportunidades, sincronizar **todos** los documentos y recargar los de la oportunidad seleccionada.
- La rutina de lectura de oportunidades tambien se usa tras altas/cambios y antes de cargar un archivo; esas invocaciones vuelven a sincronizar.
- Comparar asignacion de cada documento contra su oportunidad, considerando null = sin asignar. Solo escribir diferencias y usar ETag documental.
- Si no se encuentra la oportunidad de un documento, informar el ID y no inventar un responsable.
- Detener al primer fallo de sincronizacion e informar cantidad ya actualizada, documento pendiente y causa. La operacion es parcial, no atomica.
- Reintentar con Actualizar: volver a leer y omitir documentos que ya coinciden.
- Si la reasignacion de la oportunidad se guardo pero falla la propagacion, decirlo explicitamente; no revertir ni afirmar sincronizacion completa.
- Si otro cliente cambia el responsable en la oportunidad, actualizar la web lo propaga. Si cambia solo el documento, la siguiente sincronizacion lo sobrescribe desde la oportunidad.
- **Sin polling, push ni proceso de fondo**: con la pagina cerrada no hay sincronizacion automatica. Tampoco hay garantia transaccional entre una modificacion externa concurrente y la copia del metadato; una nueva actualizacion reconcilia diferencias.

## 7. Servidor local y entregables

Crear:

- `index.html`: aplicacion completa, configurada para el sitio destino.
- `Start-Dashboard.ps1`: servidor local sin dependencia de Node/Python.
- `tests\assignment.test.cjs`: pruebas automatizadas de miembros, herencia, paginacion y conflictos.
- `README.md`: instrucciones reales del entorno creado, esquema, autenticacion, limites, ejecucion y verificacion.

Servidor PowerShell:

- Parametro `Port`, predeterminado 3000, rango 1024 a 65535.
- `System.Net.Sockets.TcpListener` sobre `IPAddress.Loopback`.
- Servir solo GET `/` y `/index.html`, leyendo el HTML del directorio del script.
- Responder 404 para otras rutas y 405 para otros metodos, sin exponer carpetas o archivos del equipo.
- UTF-8, Content-Length correcto, `Cache-Control: no-store`, cerrar conexion; timeouts de lectura/escritura y liberacion de cliente/listener.
- Ejecutar con `.\Start-Dashboard.ps1` y abrir `http://localhost:3000/`. No usar `file://`.
- Si PowerShell bloquea el script, no cambiar politicas ni eludir restricciones: usar un servidor aprobado por la organizacion. Si se cambia puerto/ruta, actualizar la SPA de forma coherente.

No poblar el sitio con datos de demostracion. Para pruebas reales crear registros temporales claramente identificados, con autorizacion, y eliminar solo esos registros/documentos al finalizar.

## 8. Prompts para enviar a Copilot

### Prompt 0 - Contexto e unica entrada editable

```text
Quiero construir desde cero la aplicacion descrita en el archivo adjunto
GUIA-RECREAR-CON-COPILOT.md. Consideralo la especificacion completa.

SITIO_SHAREPOINT = "https://m365x30064847.sharepoint.com/sites/test-demo-ia"

Esta URL es la unica entrada de configuracion que proporciono. Descubre el
tenant y los IDs de sitio/listas/grupo con herramientas autorizadas; crea o
reutiliza un registro Entra compatible y captura su client ID como resultado
del aprovisionamiento. Configura esos valores en el proyecto tu mismo.
No reutilices identificadores del entorno original ni me entregues codigo
con placeholders pendientes como si estuviera terminado.

Trabaja en la carpeta actual. Usa HTML y JavaScript nativos, SharePoint REST,
MSAL Browser y un servidor PowerShell loopback. No agregues backend, Graph,
secretos, datos ficticios, Power Automate ni almacenamiento local de negocio.
Conserva exactamente listas, nombres internos, estados y comportamientos
definidos en la guia.

Primero inspecciona el entorno y las herramientas disponibles. Identifica
operaciones que requieren mi login, MFA, consentimiento o autorizacion
administrativa. No cambies permisos del sitio ni politicas de seguridad.
No declares aprovisionamiento o pruebas reales sin evidencia.
En este paso confirma el alcance y las condiciones para continuar, sin
modificar todavia los recursos remotos.
```

### Prompt 1 - Aprovisionar las listas y columnas

```text
Ejecuta la seccion 3 de la guia para SITIO_SHAREPOINT, con mi cuenta
autorizada. Inspecciona primero lo existente y aprovisiona idempotentemente
las dos listas y todos sus campos, tipos, opciones y propiedades.

Crea nombres internos ASCII antes de aplicar etiquetas visibles. Descubre
el GUID de Oportunidades para el lookup y el ID del grupo de miembros para
SelectionGroup en ambas columnas AsignadoA; no uses IDs fijos.
Habilita adjuntos y versiones en DocumentosOportunidades, indexa el lookup
y configura Restrict. Conserva asignaciones historicas y datos existentes.

Si tus herramientas no permiten operar autenticado, dame los pasos exactos
de portal o un script revisable para ejecutarlos con permisos autorizados;
recoge los resultados antes de continuar. No des por creado lo pendiente.
Verifica mediante lectura posterior tipos, Required, opciones, lookup,
restriccion, adjuntos, versiones y vistas. Informa discrepancias sin
destruir recursos incompatibles.
```

### Prompt 2 - Crear y verificar la autenticacion

```text
Implementa el aprovisionamiento Entra de la seccion 4: SPA single tenant,
http://localhost:3000/, SharePoint AllSites.Write delegado, sin secretos,
sin flujo implicito, sin cliente publico y sin Graph User.Read innecesario.
No concedas consentimiento global automaticamente.

Descubre y verifica el tenant destino; crea o reutiliza un registro
compatible y captura su Application (client) ID. Si necesitas pasos de
portal, guiame y recoge sus resultados: esos IDs son salidas de la
instalacion, no parametros que deba reemplazar en los prompts siguientes.

Deja documentados el registro real, las redirect URI, el permiso y el
estado de consentimiento. Distingue los permisos administrativos usados
para instalar de los permisos delegados de la SPA. No afirmes que la
autenticacion funciona hasta completar un acceso interactivo y una lectura.
```

### Prompt 3 - Construir el dashboard y CRUD de oportunidades

```text
Crea index.html y Start-Dashboard.ps1 conforme a las secciones 2, 4, 5, 6
y 7. Completa CONFIG con los valores obtenidos y deriva todos los enlaces
al sitio destino. Implementa el aspecto Microsoft 365, conexion MSAL,
restauracion de sesion y fallback popup/redirect.

Implementa KPIs con las formulas exactas, formulario validado, filtros
con cantidades, tabla/tarjetas, asignacion y cambio de estado desde ambas
vistas, eliminacion confirmada y bloqueo por documentos relacionados.
Solo ofrecer miembros reales del grupo asociado y grupos Microsoft 365
enumerables como indica la guia; revalidar y usar ensureuser al guardar.

Centraliza REST con paginacion, validacion de URL, ETag y errores visibles.
No fabriques exito ni actualices datos como guardados antes de confirmacion.
Distingue escritura correcta de refresco fallido.
No agregues edicion completa ni funcionalidades fuera de la especificacion.
Deja preparada la integracion con el panel documental del paso siguiente.
```

### Prompt 4 - Adjuntos y validacion documental

```text
Implementa el panel documental completo de las secciones 5 y 6:
seleccion de oportunidad, carga binaria como adjunto de elemento de lista,
listado filtrado, descarga autenticada, revision manual y eliminacion.

Respeta 20*1024*1024 bytes, nombres de 180 caracteres, archivos no vacios,
escape de apostrofes, observaciones de 2000 y motivo obligatorio al rechazar.
Validado/Rechazado guardan cuenta y fecha UTC; Pendiente limpia ambos.
La revision y eliminacion deben usar ETag exacto.

Lee la asignacion actual antes de crear el registro documental. Implementa
la compensacion si falla el binario y comunica el ID si no se confirma la
limpieza. Un refresco fallido tras carga correcta no debe invitar a recargar
el mismo archivo. No expongas documentos anteriores como si fueran actuales
si falla la lectura. No uses biblioteca documental, Graph ni base64 en
localStorage. No presentes la validacion manual como analisis automatico.
```

### Prompt 5 - Completar la sincronizacion de responsables

```text
Implementa y conecta la sincronizacion completa de la seccion 6. La
oportunidad es la fuente de verdad; cada documento guarda AsignadoAId
como metadato. No cambies bytes del archivo ni otras columnas al sincronizar.

Al conectar, restaurar, Actualizar y Actualizar documentos, leer miembros y
oportunidades y sincronizar todos los documentos con paginacion, no solo
los visibles. Reutiliza esa rutina tras cambios y antes de nuevas cargas.
Escribe solo diferencias, incluidos null historicos, usando el ETag exacto.

Ante fallo parcial, detente e informa cantidad aplicada, documento y error.
Reintentar debe omitir documentos ya sincronizados. Distingue una
reasignacion guardada con propagacion incompleta. Si falta la oportunidad,
no asignes un responsable por defecto. Conserva historicos fuera del grupo.
Sin tareas de fondo ni ampliacion de permisos; deja clara esa limitacion.
```

### Prompt 6 - Pruebas y verificacion real

```text
Ejecuta la matriz de la seccion 9. Agrega pruebas automatizadas enfocadas
con node:test, sin nuevas dependencias, para miembros, paginacion, ETags,
sincronizacion parcial, reintento y lectura de asignacion antes de carga.
Prueba tambien validaciones y flujos de documentos con mocks donde proceda.

Inicia el servidor aprobado, comprueba que responde y prueba la interfaz
con navegador en escritorio y a 390 px. Si estoy autorizado y tengo sesion,
crea datos temporales de prueba en SharePoint y verifica persistencia desde
la API o el sitio, no solo desde el DOM. Confirma los bytes descargados,
el conflicto 412 y la restriccion de eliminacion en el servidor.

No hagas pruebas destructivas sobre datos ajenos. Elimina unicamente los
temporales que creaste y verifica su ausencia. Separa resultados
automatizados, simulados, reales y bloqueados; no simules pruebas reales.
```

### Prompt 7 - Cierre y documentacion

```text
Completa README.md con la configuracion real obtenida, esquema, Entra,
ejecucion Windows, errores, sincronizacion y limites de esta guia.
No incluyas tokens ni secretos. Explica las restricciones de consentimiento,
roles, adjuntos y sincronizacion solo con la pagina abierta.

Verifica que no queden URLs o IDs del entorno original, placeholders
ejecutables, datos ficticios ni almacenamiento local de negocio.
Entrega los archivos persistidos y un informe breve con pruebas ejecutadas,
evidencias de SharePoint/Entra comprobadas, limpieza y pendientes reales.
Si algun permiso o paso interactivo impidio completar algo, indicalo
expresamente; no declares la replica terminada por tener solo la UI.
```

## 9. Matriz de aceptacion

| Area | Prueba y resultado requerido |
| --- | --- |
| Aprovisionamiento | Reejecutar no duplica listas/campos ni borra datos; tipos incompatibles se informan. Leer esquema real y comprobar todas las propiedades de la seccion 3. |
| Entra | Single tenant y SPA correctos, redirect exacto, solo permiso delegado requerido, sin secretos. Lectura OAuth desde localhost y recarga de sesion. Probar fallback de popup de forma controlada. |
| Servidor | `/` y `/index.html` responden 200; ruta ajena 404 y POST 405. Solo loopback, sin servir archivos arbitrarios. |
| Alta | Crear oportunidad valida, comprobar valores e ID en SharePoint, refrescar y recargar navegador; sigue existiendo. Rechazar vacios, presupuesto negativo/invalido y cuenta ajena. |
| KPIs | Con presupuestos temporales Prospecto=10, Negociación=20, Ganada=30, Perdida=40: pipeline 100, ingresos 30, abiertas 2 y tasa 50%. Sin cerradas: 0%. Los filtros no cambian KPIs. |
| Miembros | Expandir y paginar grupo Microsoft 365, admitir personas directas, deduplicar, rechazar cuenta ajena y grupos no soportados. No consultar todos los usuarios como alternativa. |
| Concurrencia | Leer ETag, modificar con otra solicitud y probar mutacion con ETag viejo: HTTP 412 sin sobrescribir. Hacerlo en oportunidades y documentos. |
| Estados/vistas | Cambiar estado y responsable en tabla y tarjetas; comprobar SharePoint, conteos y vista recordada. Si falla escritura, no mostrar guardado. |
| Paginacion | Respuestas de varias paginas para oportunidades, miembros, documentos filtrados y sincronizacion global; no perder registros de paginas posteriores. Rechazar nextLink fuera del sitio. |
| Archivo | Subir, verificar elemento + adjunto y descargar exactamente los mismos bytes. Probar apostrofe en nombre, vacio, nombre invalido, 180/181 caracteres y limite de bytes exacto/mas uno. Los limites pueden probarse con mocks sin cargas grandes reales. |
| Revision | Pendiente inicial, Validado con cuenta/fecha, Rechazado sin motivo bloqueado, con motivo guardado; volver a Pendiente limpia cuenta/fecha actual y conserva historial. |
| Aislamiento | Dos oportunidades con documentos: selector/filtro muestra solo los de la elegida. Cambiar seleccion no conserva formulario/archivo anterior. |
| Herencia | Alta y nueva carga heredan ID de persona; lectura previa detecta reasignacion mas reciente. Cambiar responsable propaga a documentos sin alterar binarios ni revision. |
| Cambios externos | Modificar oportunidad desde otro cliente y actualizar web propaga responsable. Cambiar solo el documento se sobrescribe en la siguiente sincronizacion. Pagina cerrada no sincroniza. |
| Fallo parcial | Conflicto en segundo documento informa primero aplicado y segundo pendiente; reintento no vuelve a escribir el primero. |
| Recuperacion | Simular fallo binario: limpia provisional; si limpieza falla informa ID. Simular guardado correcto + refresco fallido: informa persistencia sin sugerir duplicar. Error de lectura no muestra documentos viejos como actuales. |
| Eliminacion | Con documentos, la web y SharePoint bloquean borrar oportunidad. Borrar documento elimina adjunto y elemento pero no oportunidad. Despues se puede borrar oportunidad vacia. |
| Responsive | A 390 px sin desbordamiento de pagina, controles accesibles y scroll de tabla interno; escritorio con paneles previstos. |
| Limpieza | Confirmar eliminacion de documentos y oportunidades temporales, sin modificar datos previos ni pertenencias/permisos del sitio. |

Las cinco pruebas automatizadas del proyecto original cubren: expansion/paginacion de miembros y rechazo de ajenos; error de grupo no soportado; sincronizacion de diferencias con ETags y paginacion; fallo parcial/reintento; asignacion actualizada antes de la carga. La matriz anterior tambien especifica comprobaciones adicionales necesarias para validar la replica completa.

## 10. Que se replica y que no se debe trasladar del original

El repositorio original documenta un registro Entra llamado `demo-ia-with-copilot-dashboard`, SPA localhost, SharePoint `AllSites.Write` delegado, sin secreto, sin flujos implicitos ni cliente publico, sin Graph `User.Read` y sin consentimiento global al directorio; el primer acceso se documento con consentimiento para la cuenta conectada. Eso describe el entorno original, **no garantiza** la misma politica de consentimiento en otro tenant.

Se reproducen el modelo, la UX y las operaciones, no los datos de negocio, cuentas, IDs de elementos, ETags, GUIDs de tenant/cliente/lista ni IDs de grupos del sitio anterior. Tampoco hay que recrear sus cinco oportunidades/documentos historicos: una instalacion nueva puede empezar vacia.

Fuente de esta guia: inspeccion de `index.html`, `Start-Dashboard.ps1`, `tests\assignment.test.cjs` y `README.md` del proyecto. No se consulto el portal de Entra ni el sitio en vivo al escribirla. Las verificaciones historicas descritas por el README no sustituyen las pruebas en el sitio destino.
