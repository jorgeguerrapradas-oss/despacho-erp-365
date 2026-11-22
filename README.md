# ERP del despacho sobre Microsoft 365

Repositorio **privado** con la configuración, plantillas y diseños para automatizar la gestión del despacho (laboral, fiscal, contable, civil, mercantil y herencias) sobre **CEGID Despachos + Microsoft 365** (SharePoint, Power Automate, Copilot) con apoyo de Claude.

> Este repositorio contiene **solo configuración y documentación interna**. No se suben datos de clientes, expedientes reales ni credenciales.

## Estructura

| Carpeta | Contenido | Estado |
|---|---|---|
| `plantillas/` | `Plantilla_Expedientes_SharePoint.xlsx`: plantilla de expedientes (11 hojas, checklists por área con base legal). Códigos de área: HER, LAB, FIS, CIV, MER, CON, ADM, OTR. | v1 |
| `sharepoint-script/` | Scripts PnP.PowerShell para crear (`Crear-ListasExpedientes.ps1`) y revertir (`Revertir-ListasExpedientes.ps1`) la biblioteca **Expedientes** y la lista **Control de Trámites**, con `INSTRUCCIONES.md` paso a paso. | Sin ejecutar aún |
| `automatizaciones/` | `Diseno_automatizaciones_Power_Automate.docx`: diseño de 4 flujos de Power Automate. | Diseño |
| `docs/` | Documentación adicional (decisiones, procedimientos). | Vacía |

### Nombres internos que no deben cambiarse

Los flujos de Power Automate dependen de estos nombres internos de columna. Si se cambian, hay que actualizar a la vez el script, la plantilla y los flujos.

- Biblioteca **Expedientes**: `NumExpediente`, `AreaExpediente`, `EstadoExpediente`, `ResponsableExp`, `CodigoCEGID`, `ConservarHasta`.
- Lista **Control de Trámites** (`Lists/ControlTramites`): `FechaLimite`, `EstadoTramite`, `ResponsableTram`, `AvisoEnviado`.

## Por qué es privado y qué no se sube (base legal)

- **Seguridad y confidencialidad de los datos personales**: art. 5.1.f) y art. 32 del Reglamento (UE) 2016/679 (RGPD), y art. 5 de la LO 3/2018 (LOPDGDD), deber de confidencialidad.
- **Secreto profesional**: art. 542.3 de la LO 6/1985 del Poder Judicial, para los asuntos con intervención de abogado.
- **GitHub como proveedor**: si algún día se subieran datos personales, GitHub actuaría como encargado del tratamiento (art. 28 RGPD) y habría transferencia internacional (arts. 44 a 46 RGPD). Por eso la regla es **no subir datos personales**: el `.gitignore` excluye CSV, exportaciones y certificados.

## Próximos pasos

1. **Ejecutar el script de SharePoint en un sitio de pruebas**, siguiendo `sharepoint-script/INSTRUCCIONES.md` (primero con `-WhatIf`).
2. **Construir los 4 flujos de Power Automate** del documento de diseño, empezando por el aviso de plazos de la lista Control de Trámites.
3. **Conectar CEGID Despachos** con SharePoint mediante el campo `CodigoCEGID` (exportación/importación o API, según licencia).
4. **Configurar Copilot** sobre la biblioteca Expedientes (permisos por área antes de activarlo).
5. **Documentar en `docs/`** cada decisión y el registro de actividades de tratamiento afectado (art. 30 RGPD).

## Cómo trabajar con el repositorio

- Cada cambio en script, plantilla o flujo, en una rama propia y con *pull request*, para que quede historial de quién cambió qué.
- Exportar los flujos de Power Automate como paquete (.zip) a `automatizaciones/` cuando estén construidos, revisando antes que no contengan datos de clientes.
