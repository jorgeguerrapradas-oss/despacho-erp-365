# Crear la biblioteca de Expedientes y la lista de Control de Trámites en SharePoint

Esta carpeta contiene dos scripts de PowerShell:

| Archivo | Qué hace |
|---|---|
| `Crear-ListasExpedientes.ps1` | Crea la biblioteca **Expedientes** y la lista **Control de Trámites**, con sus columnas, índices y vistas. Se puede ejecutar varias veces sin duplicar nada. |
| `Revertir-ListasExpedientes.ps1` | Deshace lo anterior. Envía la biblioteca y la lista a la papelera de reciclaje y se niega a borrarlas si ya tienen documentos. |

Tiempo estimado: 20 minutos la primera vez (pasos 1 a 3 solo se hacen una vez) y 2 minutos las siguientes.

---

## Antes de empezar

- **Quién**: necesita ser **propietario del sitio** de SharePoint donde se crearán las listas. Para el paso 3 (registrar la aplicación) necesita además ser **administrador global** o **administrador de aplicaciones** del tenant de Microsoft 365, o pedírselo a quien lo sea.
- **Dónde**: un equipo con Windows 10 u 11.
- **Sitio**: cree antes el sitio de SharePoint (por ejemplo, un sitio de grupo llamado *Expedientes*) y copie su URL, por ejemplo `https://midespacho.sharepoint.com/sites/Expedientes`.
- **Recomendación**: haga la primera prueba en un **sitio de pruebas**, no en el sitio definitivo.

---

## Paso 1. Instalar PowerShell 7

La versión actual de PnP.PowerShell (3.x) exige **PowerShell 7.4 o superior**. El "Windows PowerShell 5.1" que viene con Windows no sirve.

1. Pulse **Inicio**, escriba `cmd` y abra **Símbolo del sistema**.
2. Ejecute:
   ```
   winget install --id Microsoft.PowerShell --source winget
   ```
3. Cierre la ventana. A partir de ahora abra **PowerShell 7** desde el menú Inicio (icono negro, se llama "PowerShell 7 (x64)").
4. Compruebe la versión:
   ```powershell
   $PSVersionTable.PSVersion
   ```
   Debe ser 7.4 o mayor.

## Paso 2. Instalar el módulo PnP.PowerShell

En PowerShell 7:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Si pregunta si confía en el repositorio "PSGallery", responda `S` (Sí). Para comprobarlo:

```powershell
Get-Module PnP.PowerShell -ListAvailable | Select-Object Name, Version
```

## Paso 3. Registrar la aplicación de PnP en Entra ID (solo una vez por despacho)

Desde septiembre de 2024 PnP PowerShell ya no trae una aplicación compartida: cada organización debe registrar la suya en Entra ID (antes Azure AD) y usar su *Client ID* al conectarse.

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell Despacho" -Tenant midespacho.onmicrosoft.com -Interactive
```

- Cambie `midespacho.onmicrosoft.com` por el dominio `onmicrosoft.com` de su despacho (aparece en el Centro de administración de Microsoft 365 > Configuración > Dominios).
- Se abrirá el navegador: inicie sesión con la cuenta de administrador y pulse **Aceptar** en la pantalla de permisos.
- Al terminar, la consola muestra un **AzureAppId / ClientId** (un código tipo `1a2b3c4d-...`). **Cópielo y guárdelo**: lo usará siempre.

## Paso 4. Copiar los scripts al equipo

1. Cree la carpeta `C:\Scripts\SharePoint`.
2. Copie dentro `Crear-ListasExpedientes.ps1` y `Revertir-ListasExpedientes.ps1`.
3. Windows marca como "bloqueados" los archivos descargados. Desbloquéelos:
   ```powershell
   cd C:\Scripts\SharePoint
   Get-ChildItem *.ps1 | Unblock-File
   ```
4. Si al ejecutar aparece un error de "ejecución de scripts deshabilitada", permita scripts locales solo para su usuario:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   ```

## Paso 5. (Opcional) Ajustar columnas

Abra `Crear-ListasExpedientes.ps1` con el Bloc de notas o VS Code. Todo lo configurable está en la sección **CONFIGURACIÓN** al principio:

- `$ColumnasExpedientes`: columnas de la biblioteca (hoja *Columnas Comunes*).
- `$ColumnasTramites`: columnas de la lista (hoja *Control de Trámites*).
- `$VistasExpedientes` y `$VistasTramites`: vistas.

Reglas: el **nombre interno** (`Interno`) no lleva tildes, espacios ni ñ, y no se puede cambiar una vez creada la columna. El **nombre visible** (`Nombre`) sí admite tildes y se puede cambiar después desde SharePoint.

## Paso 6. Simulación (no cambia nada)

```powershell
cd C:\Scripts\SharePoint
.\Crear-ListasExpedientes.ps1 -SiteUrl "https://midespacho.sharepoint.com/sites/Expedientes" -ClientId "SU-CLIENT-ID" -WhatIf
```

Se abrirá el navegador para iniciar sesión. La consola listará líneas `What if: ...` con todo lo que se crearía. No se crea nada.

## Paso 7. Ejecución real

```powershell
.\Crear-ListasExpedientes.ps1 -SiteUrl "https://midespacho.sharepoint.com/sites/Expedientes" -ClientId "SU-CLIENT-ID"
```

Verá en verde (`+`) lo creado y en gris (`=`) lo que ya existía. En la misma carpeta queda un registro `Crear-ListasExpedientes_AAAAMMDD_HHMMSS.log`.

Si algo falla a mitad, el script se para sin borrar nada. Corrija la causa y vuelva a ejecutarlo: continúa donde se quedó.

## Paso 8. Comprobar en SharePoint

1. Abra el sitio. En el menú de la izquierda aparecerán **Expedientes** y **Control de Trámites**.
2. En **Control de Trámites**, cree un trámite de prueba con fecha límite dentro de 10 días: debe aparecer en la vista *Plazos próximos (15 días)*.
3. En **Expedientes**, suba un documento, rellene sus propiedades (Nº Expediente, Área…) y compruebe las vistas *Por área* y *Mis expedientes*.

---

## Qué se crea

### Biblioteca *Expedientes* (hoja *Columnas Comunes*)

Control de versiones activado (100 versiones) y creación de carpetas permitida.

| Columna | Tipo | Notas |
|---|---|---|
| Nº Expediente | Texto | Indexada. Formato según la hoja *Numeración Expedientes* |
| Cliente | Texto | Indexada |
| NIF/CIF cliente | Texto | Indexada |
| Código cliente CEGID | Texto | Indexada. Enlace con el ERP |
| Área | Elección | Laboral, Fiscal, Contable, Civil, Mercantil, Herencias y sucesiones, Administrativo, Otros |
| Tipo de documento | Elección | Escrito, Contrato, Modelo tributario, Notificación… |
| Estado del expediente | Elección | Abierto (por defecto), En trámite, Pendiente del cliente, Cerrado, Archivado |
| Responsable | Persona | Indexada |
| Ejercicio | Texto | Indexada |
| Fecha de apertura / Fecha de cierre | Fecha | Solo fecha |
| Confidencialidad | Elección | Sin datos personales, Datos personales, Categorías especiales (art. 9 RGPD) |
| Conservar hasta | Fecha | Fin del plazo de conservación |
| Observaciones | Texto largo | |

En la biblioteca **ninguna columna es obligatoria a propósito**: con columnas obligatorias, los archivos que se suben o sincronizan con OneDrive quedan marcados como "falta información necesaria". La obligatoriedad se controlará mejor con el flujo de alta de expedientes en Power Automate.

Vistas: *Todos los expedientes* (predeterminada), *Por área*, *Mis expedientes*, *Pendientes de purga RGPD*.

### Lista *Control de Trámites* (hoja *Control de Trámites*)

Control de versiones activado (50 versiones). La columna *Título* pasa a llamarse **Trámite**.

| Columna | Tipo | Notas |
|---|---|---|
| Trámite | Texto | Obligatoria (columna Título) |
| Nº Expediente | Texto | Obligatoria, indexada |
| Cliente | Texto | Indexada |
| Área | Elección | Mismas áreas que la biblioteca |
| Organismo | Elección | AEAT, TGSS, SEPE, INSS, Juzgado / Tribunal, SMAC / Mediación, Registros, Notaría, CC. AA., Ayuntamiento, Otro |
| Fecha notificación / inicio | Fecha | Desde cuándo cuenta el plazo |
| Fecha límite | Fecha | Obligatoria, indexada |
| Cómputo del plazo | Elección | Días hábiles, Días naturales, Meses, Años |
| Base legal | Texto | Norma y artículo del trámite |
| Estado | Elección | Pendiente (por defecto), En curso, Presentado, Cerrado, Desistido |
| Prioridad | Elección | Alta, Normal (por defecto), Baja |
| Responsable | Persona | Indexada |
| Fecha de presentación | Fecha | |
| Nº registro / justificante | Texto | |
| Aviso enviado | Sí/No | Lo usará Power Automate al enviar avisos de vencimiento |
| Observaciones | Texto largo | |

Vistas: *Plazos próximos (15 días)* (predeterminada), *Vencidos sin presentar*, *Mis trámites*, *Por expediente*, *Todos los trámites*. Las tres primeras ocultan los trámites Presentados, Cerrados o Desistidos.

---

## Cómo revertirlo

### Simulación

```powershell
.\Revertir-ListasExpedientes.ps1 -SiteUrl "https://midespacho.sharepoint.com/sites/Expedientes" -ClientId "SU-CLIENT-ID" -WhatIf
```

### Borrar biblioteca y lista (solo si están vacías)

```powershell
.\Revertir-ListasExpedientes.ps1 -SiteUrl "https://midespacho.sharepoint.com/sites/Expedientes" -ClientId "SU-CLIENT-ID"
```

Pide confirmación (`S`) antes de cada borrado. Si la biblioteca o la lista tienen contenido, **no las borra** y avisa. Para borrarlas aun así, añada `-IncluirConContenido`.

### Quitar solo columnas y vistas, conservando los documentos

```powershell
.\Revertir-ListasExpedientes.ps1 -SiteUrl "..." -ClientId "SU-CLIENT-ID" -SoloVistasYColumnas
```

Atención: al quitar una columna se pierden los valores guardados en ella. Los documentos se conservan.

### Recuperar algo borrado

La biblioteca y la lista van a la **papelera de reciclaje** del sitio, no se destruyen. Para recuperarlas: engranaje ⚙️ > **Contenido del sitio** > **Papelera de reciclaje** > seleccionar > **Restaurar**. SharePoint Online conserva lo eliminado **93 días** entre la papelera del sitio y la papelera de segunda fase.

---

## Base legal de las decisiones de diseño

| Decisión | Fundamento |
|---|---|
| Columna *Confidencialidad* y vista de documentos de categorías especiales | Art. 9 del Reglamento (UE) 2016/679 (RGPD): datos de salud, afiliación sindical, etc., aparecen a menudo en expedientes laborales y de incapacidad. Exigen medidas reforzadas. |
| Columna *Conservar hasta* y vista *Pendientes de purga RGPD* | Art. 5.1.e RGPD (limitación del plazo de conservación) y art. 32 LOPDGDD (Ley Orgánica 3/2018), que obliga a bloquear los datos en lugar de borrarlos mientras puedan exigirse responsabilidades. |
| Referencias habituales para fijar *Conservar hasta* | 4 años de prescripción tributaria (art. 66 Ley 58/2003, LGT); 6 años para libros y documentación mercantil (art. 30 Código de Comercio); 10 años para documentación de prevención del blanqueo, a la que están obligados los asesores fiscales y contables (arts. 2.1.m y 25 Ley 10/2010). |
| Control de versiones y registro de ejecución | Art. 32 RGPD (seguridad del tratamiento: integridad, trazabilidad y capacidad de restaurar). |
| Columna *Cómputo del plazo* | Los plazos se cuentan de forma distinta según el ámbito: procedimiento administrativo, art. 30 Ley 39/2015 (días hábiles, excluidos sábados, domingos y festivos); procesal civil, arts. 130 y 133 LEC; agosto inhábil en vía judicial, art. 183 LOPJ, con las excepciones de cada orden (p. ej. art. 43.4 LRJS en lo social). |

Estas referencias son para orientar el diseño. Los plazos concretos de cada trámite deben comprobarse en la norma aplicable en cada caso.

---

## Fuentes técnicas

- PnP PowerShell, registro de la aplicación de Entra ID: <https://pnp.github.io/powershell/articles/registerapplication.html>
- Cmdlet `Register-PnPEntraIDAppForInteractiveLogin`: <https://pnp.github.io/powershell/cmdlets/Register-PnPEntraIDAppForInteractiveLogin.html>
- Cambios de PnP PowerShell v3 (requisito de PowerShell 7.4): <https://pnp.github.io/blog/pnp-powershell/pnp-powershell-v3-0-0-preview>
- Retirada de la aplicación compartida de PnP (septiembre de 2024): <https://office365itpros.com/2024/08/29/pnp-powershell-changes-app/>

## Problemas frecuentes

| Mensaje | Solución |
|---|---|
| `The term 'Connect-PnPOnline' is not recognized` | Está en Windows PowerShell 5.1 en lugar de PowerShell 7, o falta el paso 2. |
| `AADSTS700016: Application with identifier ... was not found` | El `-ClientId` es incorrecto o falta el paso 3. |
| `Access denied` | Su cuenta no es propietaria del sitio. |
| `running scripts is disabled on this system` | Ejecute el último comando del paso 4. |
| Una columna da error al crearse | Revise que el nombre interno no tenga tildes, espacios ni ñ, y vuelva a ejecutar. |
