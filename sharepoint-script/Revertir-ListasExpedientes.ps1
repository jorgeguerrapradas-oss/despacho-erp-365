<#
.SYNOPSIS
    Deshace lo creado por Crear-ListasExpedientes.ps1: elimina la biblioteca "Expedientes"
    y la lista "Control de Trámites" (con sus columnas y vistas).

.DESCRIPTION
    - Envía la biblioteca y la lista a la PAPELERA DE RECICLAJE del sitio (no las borra
      definitivamente). Se pueden restaurar desde Configuración del sitio > Papelera de
      reciclaje durante 93 días (primera fase + segunda fase de la papelera de SharePoint Online).
    - SEGURIDAD: si la biblioteca o la lista contienen documentos o elementos, NO las borra
      salvo que se añada -IncluirConContenido. Así no se pierden expedientes por error.
    - Pide confirmación antes de cada borrado (salvo que se use -Confirm:$false).
    - Admite -WhatIf para ver qué haría sin borrar nada.

.PARAMETER SiteUrl
    URL del sitio, la misma usada al crear.

.PARAMETER ClientId
    Id. de aplicación (cliente) de la app de Entra ID para PnP PowerShell.

.PARAMETER IncluirConContenido
    Permite borrar la biblioteca/lista aunque tenga contenido. Úselo solo si está seguro.

.PARAMETER SoloVistasYColumnas
    No borra la biblioteca ni la lista: solo quita las columnas y vistas creadas por el script.
    Útil si ya hay documentos dentro y solo quiere rehacer la estructura.
    ATENCIÓN: al borrar una columna se pierden los valores guardados en ella.

.EXAMPLE
    .\Revertir-ListasExpedientes.ps1 -SiteUrl https://midespacho.sharepoint.com/sites/Expedientes -ClientId <id> -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)][string]$SiteUrl,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [switch]$IncluirConContenido,
    [switch]$SoloVistasYColumnas
)

$ErrorActionPreference = 'Stop'

# Deben coincidir con Crear-ListasExpedientes.ps1
$NombreBiblioteca = 'Expedientes'
$NombreLista      = 'Control de Trámites'

$CamposBiblioteca = @('NumExpediente', 'Cliente', 'NIFCliente', 'CodigoCEGID', 'AreaExpediente', 'TipoDocumento',
                      'EstadoExpediente', 'ResponsableExp', 'Ejercicio', 'FechaApertura', 'FechaCierre',
                      'Confidencialidad', 'ConservarHasta', 'ObservacionesExp')
$VistasBiblioteca = @('Todos los expedientes', 'Por área', 'Mis expedientes', 'Pendientes de purga RGPD')

$CamposLista = @('NumExpediente', 'Cliente', 'AreaTramite', 'Organismo', 'FechaNotificacion', 'FechaLimite',
                 'TipoComputo', 'BaseLegal', 'EstadoTramite', 'Prioridad', 'ResponsableTram', 'FechaPresentacion',
                 'NumJustificante', 'AvisoEnviado', 'ObservacionesTram')
$VistasLista = @('Plazos próximos (15 días)', 'Vencidos sin presentar', 'Mis trámites', 'Por expediente', 'Todos los trámites')

function Remove-Estructura {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string]$Lista, [string[]]$Campos, [string[]]$Vistas)

    # Antes de borrar vistas, deja como predeterminada la vista estándar (no se puede borrar la vista por defecto)
    $vistaBase = Get-PnPView -List $Lista | Where-Object { $Vistas -notcontains $_.Title -and -not $_.Hidden } | Select-Object -First 1
    if ($vistaBase -and $PSCmdlet.ShouldProcess($Lista, "Poner '$($vistaBase.Title)' como vista predeterminada")) {
        $null = Set-PnPView -List $Lista -Identity $vistaBase.Id -Values @{ DefaultView = $true }
    }
    foreach ($v in $Vistas) {
        if ((Get-PnPView -List $Lista | Select-Object -ExpandProperty Title) -contains $v -and
            $PSCmdlet.ShouldProcess($Lista, "Eliminar vista '$v'")) {
            Remove-PnPView -List $Lista -Identity $v -Force
            Write-Host "   - Vista '$v' eliminada." -ForegroundColor Yellow
        }
    }
    foreach ($f in $Campos) {
        if (Get-PnPField -List $Lista -Identity $f -ErrorAction SilentlyContinue) {
            if ($PSCmdlet.ShouldProcess($Lista, "Eliminar columna '$f' (se pierden sus valores)")) {
                Remove-PnPField -List $Lista -Identity $f -Force
                Write-Host "   - Columna '$f' eliminada." -ForegroundColor Yellow
            }
        }
    }
}

function Remove-ListaSegura {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string]$Lista)

    $l = Get-PnPList -Identity $Lista -ErrorAction SilentlyContinue
    if (-not $l) { Write-Host "   = '$Lista' no existe; nada que hacer." -ForegroundColor DarkGray; return }

    if ($l.ItemCount -gt 0 -and -not $IncluirConContenido) {
        Write-Host "   ! '$Lista' contiene $($l.ItemCount) elemento(s). NO se borra." -ForegroundColor Red
        Write-Host "     Si de verdad quiere borrarla, repita con -IncluirConContenido." -ForegroundColor Red
        return
    }
    if ($PSCmdlet.ShouldProcess($Lista, "Enviar a la papelera de reciclaje ($($l.ItemCount) elementos)")) {
        $null = Remove-PnPList -Identity $Lista -Recycle -Force
        Write-Host "   - '$Lista' enviada a la papelera de reciclaje." -ForegroundColor Yellow
    }
}

$log = Join-Path $PSScriptRoot ("Revertir-ListasExpedientes_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $log -WhatIf:$false | Out-Null
try {
    Write-Host "Conectando a $SiteUrl ..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

    if ($SoloVistasYColumnas) {
        foreach ($par in @(@($NombreBiblioteca, $CamposBiblioteca, $VistasBiblioteca), @($NombreLista, $CamposLista, $VistasLista))) {
            Write-Host "`n'$($par[0])'" -ForegroundColor Cyan
            if (Get-PnPList -Identity $par[0] -ErrorAction SilentlyContinue) {
                Remove-Estructura -Lista $par[0] -Campos $par[1] -Vistas $par[2]
            } else { Write-Host "   = No existe." -ForegroundColor DarkGray }
        }
    }
    else {
        Write-Host "`nLista '$NombreLista'" -ForegroundColor Cyan
        Remove-ListaSegura -Lista $NombreLista
        Write-Host "`nBiblioteca '$NombreBiblioteca'" -ForegroundColor Cyan
        Remove-ListaSegura -Lista $NombreBiblioteca
    }
    Write-Host "`nTerminado. Registro guardado en: $log" -ForegroundColor Cyan
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}
