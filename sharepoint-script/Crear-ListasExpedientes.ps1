<#
.SYNOPSIS
    Crea en SharePoint Online la biblioteca "Expedientes" y la lista "Control de Trámites"
    del despacho, con sus columnas, índices y vistas.

.DESCRIPTION
    Script basado en las hojas 'Columnas Comunes' y 'Control de Trámites' de
    Plantilla_Expedientes_SharePoint.xlsx.

    - Es IDEMPOTENTE: si la biblioteca, la lista, una columna o una vista ya existen,
      las salta y sigue. Se puede ejecutar varias veces sin duplicar nada.
    - Crea las columnas A NIVEL DE LISTA (no columnas de sitio). Así, al borrar la lista
      con Revertir-ListasExpedientes.ps1 desaparecen también sus columnas.
    - No borra ni modifica datos existentes.
    - Admite -WhatIf para ver qué haría sin cambiar nada.

    Para cambiar columnas, edite SOLO las tablas $ColumnasExpedientes y $ColumnasTramites
    (sección "CONFIGURACIÓN"). Los nombres internos (Interno) no deben llevar tildes,
    espacios ni la letra ñ, y NO pueden cambiarse después de crear la columna.

.PARAMETER SiteUrl
    URL del sitio de SharePoint, p. ej. https://midespacho.sharepoint.com/sites/Expedientes

.PARAMETER ClientId
    Id. de aplicación (cliente) de la app de Entra ID registrada para PnP PowerShell
    (ver INSTRUCCIONES.md, paso 3).

.EXAMPLE
    .\Crear-ListasExpedientes.ps1 -SiteUrl https://midespacho.sharepoint.com/sites/Expedientes -ClientId 00000000-0000-0000-0000-000000000000 -WhatIf

.NOTES
    Requisitos: PowerShell 7.4 o superior y módulo PnP.PowerShell 3.x.
    Permisos: propietario del sitio (o administrador de SharePoint).
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$SiteUrl,
    [Parameter(Mandatory = $true)][string]$ClientId
)

$ErrorActionPreference = 'Stop'

# =====================================================================================
#  CONFIGURACIÓN
# =====================================================================================

$NombreBiblioteca = 'Expedientes'          # Título visible de la biblioteca
$UrlBiblioteca    = 'Expedientes'          # Parte de la URL: /sites/xxx/Expedientes
$NombreLista      = 'Control de Trámites'  # Título visible de la lista
$UrlLista         = 'Lists/ControlTramites'

# Valores reutilizados en ambas listas
$Areas = @('Laboral', 'Fiscal', 'Contable', 'Civil', 'Mercantil', 'Herencias y sucesiones', 'Administrativo', 'Otros')

# Tipos admitidos en las tablas de columnas:
#   Texto, TextoLargo, Eleccion, EleccionMultiple, Fecha, Numero, Moneda, SiNo, Persona
# Claves: Interno (nombre interno, sin tildes ni espacios), Nombre (visible), Tipo,
#   Opciones (solo Eleccion), Requerida, Indexada, Defecto, Descripcion

# --- Hoja 'Columnas Comunes'  ->  biblioteca Expedientes ------------------------------
#  OJO: en la biblioteca NO se marcan columnas como obligatorias. Con columnas obligatorias,
#  los archivos subidos o sincronizados con OneDrive quedan marcados con "Falta información
#  necesaria" y pueden dar problemas de sincronización. La obligatoriedad se controla mejor
#  con Power Automate o con el formulario de alta.
$ColumnasExpedientes = @(
    @{ Interno = 'NumExpediente';     Nombre = 'Nº Expediente';        Tipo = 'Texto';    Indexada = $true
       Descripcion = 'Formato AAAA-AREA-NNNN, según la hoja Numeración Expedientes' }
    @{ Interno = 'Cliente';           Nombre = 'Cliente';              Tipo = 'Texto';    Indexada = $true }
    @{ Interno = 'NIFCliente';        Nombre = 'NIF/CIF cliente';      Tipo = 'Texto';    Indexada = $true }
    @{ Interno = 'CodigoCEGID';       Nombre = 'Código cliente CEGID'; Tipo = 'Texto';    Indexada = $true
       Descripcion = 'Código del cliente en CEGID Despachos, para enlazar con el ERP' }
    @{ Interno = 'AreaExpediente';    Nombre = 'Área';                 Tipo = 'Eleccion'; Opciones = $Areas; Indexada = $true }
    @{ Interno = 'TipoDocumento';     Nombre = 'Tipo de documento';    Tipo = 'Eleccion'
       Opciones = @('Escrito', 'Contrato', 'Modelo tributario', 'Notificación', 'Resolución', 'Sentencia', 'Poder', 'Escritura', 'Factura', 'Nómina', 'Correspondencia', 'Documentación del cliente', 'Otros') }
    @{ Interno = 'EstadoExpediente';  Nombre = 'Estado del expediente'; Tipo = 'Eleccion'; Indexada = $true; Defecto = 'Abierto'
       Opciones = @('Abierto', 'En trámite', 'Pendiente del cliente', 'Cerrado', 'Archivado') }
    @{ Interno = 'ResponsableExp';    Nombre = 'Responsable';          Tipo = 'Persona';  Indexada = $true }
    @{ Interno = 'Ejercicio';         Nombre = 'Ejercicio';            Tipo = 'Texto';    Indexada = $true
       Descripcion = 'Año fiscal o ejercicio al que se refiere el documento (p. ej. 2026)' }
    @{ Interno = 'FechaApertura';     Nombre = 'Fecha de apertura';    Tipo = 'Fecha' }
    @{ Interno = 'FechaCierre';       Nombre = 'Fecha de cierre';      Tipo = 'Fecha' }
    @{ Interno = 'Confidencialidad';  Nombre = 'Confidencialidad';     Tipo = 'Eleccion'; Defecto = 'Datos personales'
       Opciones = @('Sin datos personales', 'Datos personales', 'Categorías especiales (art. 9 RGPD)')
       Descripcion = 'Clasificación a efectos del RGPD (Reglamento UE 2016/679)' }
    @{ Interno = 'ConservarHasta';    Nombre = 'Conservar hasta';      Tipo = 'Fecha'
       Descripcion = 'Fin del plazo de conservación (art. 5.1.e RGPD: limitación del plazo de conservación)' }
    @{ Interno = 'ObservacionesExp';  Nombre = 'Observaciones';        Tipo = 'TextoLargo' }
)

# --- Hoja 'Control de Trámites'  ->  lista Control de Trámites -------------------------
#  La columna Título (Title) de la lista se renombra a "Trámite".
$ColumnasTramites = @(
    @{ Interno = 'NumExpediente';     Nombre = 'Nº Expediente';        Tipo = 'Texto';    Requerida = $true; Indexada = $true }
    @{ Interno = 'Cliente';           Nombre = 'Cliente';              Tipo = 'Texto';    Indexada = $true }
    @{ Interno = 'AreaTramite';       Nombre = 'Área';                 Tipo = 'Eleccion'; Opciones = $Areas; Indexada = $true }
    @{ Interno = 'Organismo';         Nombre = 'Organismo';            Tipo = 'Eleccion'
       Opciones = @('AEAT', 'TGSS', 'SEPE', 'INSS', 'Juzgado / Tribunal', 'SMAC / Mediación', 'Registro Mercantil', 'Registro de la Propiedad', 'Notaría', 'Comunidad Autónoma', 'Ayuntamiento', 'Otro') }
    @{ Interno = 'FechaNotificacion'; Nombre = 'Fecha notificación / inicio'; Tipo = 'Fecha'
       Descripcion = 'Fecha desde la que empieza a contar el plazo' }
    @{ Interno = 'FechaLimite';       Nombre = 'Fecha límite';         Tipo = 'Fecha';    Requerida = $true; Indexada = $true }
    @{ Interno = 'TipoComputo';       Nombre = 'Cómputo del plazo';    Tipo = 'Eleccion'
       Opciones = @('Días hábiles', 'Días naturales', 'Meses', 'Años')
       Descripcion = 'Administrativo: art. 30 Ley 39/2015. Judicial: arts. 130-136 LEC y art. 182-183 LOPJ (agosto inhábil salvo excepciones).' }
    @{ Interno = 'BaseLegal';         Nombre = 'Base legal';           Tipo = 'Texto'
       Descripcion = 'Norma y artículo que fija el trámite o plazo (p. ej. art. 59 ET, art. 223 LGT)' }
    @{ Interno = 'EstadoTramite';     Nombre = 'Estado';               Tipo = 'Eleccion'; Indexada = $true; Defecto = 'Pendiente'
       Opciones = @('Pendiente', 'En curso', 'Presentado', 'Cerrado', 'Desistido') }
    @{ Interno = 'Prioridad';         Nombre = 'Prioridad';            Tipo = 'Eleccion'; Defecto = 'Normal'
       Opciones = @('Alta', 'Normal', 'Baja') }
    @{ Interno = 'ResponsableTram';   Nombre = 'Responsable';          Tipo = 'Persona';  Indexada = $true }
    @{ Interno = 'FechaPresentacion'; Nombre = 'Fecha de presentación'; Tipo = 'Fecha' }
    @{ Interno = 'NumJustificante';   Nombre = 'Nº registro / justificante'; Tipo = 'Texto' }
    @{ Interno = 'AvisoEnviado';      Nombre = 'Aviso enviado';        Tipo = 'SiNo';     Defecto = '0'
       Descripcion = 'Lo marca Power Automate al enviar el aviso de vencimiento' }
    @{ Interno = 'ObservacionesTram'; Nombre = 'Observaciones';        Tipo = 'TextoLargo' }
)

# --- Vistas --------------------------------------------------------------------------
#  Query = CAML sin la etiqueta <Query>.  <Today OffsetDays='15'/> = hoy + 15 días.
$Abiertos = "<Neq><FieldRef Name='EstadoExpediente'/><Value Type='Choice'>Archivado</Value></Neq>"
$NoTerminado = "<And><Neq><FieldRef Name='EstadoTramite'/><Value Type='Choice'>Presentado</Value></Neq>" +
               "<And><Neq><FieldRef Name='EstadoTramite'/><Value Type='Choice'>Cerrado</Value></Neq>" +
               "<Neq><FieldRef Name='EstadoTramite'/><Value Type='Choice'>Desistido</Value></Neq></And></And>"

$VistasExpedientes = @(
    @{ Titulo = 'Todos los expedientes'; Defecto = $true
       Campos = @('DocIcon', 'LinkFilename', 'NumExpediente', 'Cliente', 'AreaExpediente', 'TipoDocumento', 'EstadoExpediente', 'ResponsableExp', 'Modified')
       Query  = "<OrderBy><FieldRef Name='NumExpediente' Ascending='FALSE'/></OrderBy>" }
    @{ Titulo = 'Por área'
       Campos = @('DocIcon', 'LinkFilename', 'NumExpediente', 'Cliente', 'TipoDocumento', 'EstadoExpediente', 'ResponsableExp')
       Query  = "<GroupBy Collapse='TRUE'><FieldRef Name='AreaExpediente'/></GroupBy><OrderBy><FieldRef Name='NumExpediente'/></OrderBy>" }
    @{ Titulo = 'Mis expedientes'
       Campos = @('DocIcon', 'LinkFilename', 'NumExpediente', 'Cliente', 'AreaExpediente', 'EstadoExpediente', 'Modified')
       Query  = "<Where><And><Eq><FieldRef Name='ResponsableExp'/><Value Type='Integer'><UserID Type='Integer'/></Value></Eq>$Abiertos</And></Where><OrderBy><FieldRef Name='Modified' Ascending='FALSE'/></OrderBy>" }
    @{ Titulo = 'Pendientes de purga RGPD'
       Campos = @('DocIcon', 'LinkFilename', 'NumExpediente', 'Cliente', 'Confidencialidad', 'ConservarHasta')
       Query  = "<Where><Leq><FieldRef Name='ConservarHasta'/><Value Type='DateTime'><Today/></Value></Leq></Where><OrderBy><FieldRef Name='ConservarHasta'/></OrderBy>" }
)

$VistasTramites = @(
    @{ Titulo = 'Plazos próximos (15 días)'; Defecto = $true
       Campos = @('LinkTitle', 'NumExpediente', 'Cliente', 'Organismo', 'FechaLimite', 'TipoComputo', 'EstadoTramite', 'Prioridad', 'ResponsableTram')
       Query  = "<Where><And><Leq><FieldRef Name='FechaLimite'/><Value Type='DateTime'><Today OffsetDays='15'/></Value></Leq>$NoTerminado</And></Where><OrderBy><FieldRef Name='FechaLimite'/></OrderBy>" }
    @{ Titulo = 'Vencidos sin presentar'
       Campos = @('LinkTitle', 'NumExpediente', 'Cliente', 'Organismo', 'FechaLimite', 'EstadoTramite', 'ResponsableTram')
       Query  = "<Where><And><Lt><FieldRef Name='FechaLimite'/><Value Type='DateTime'><Today/></Value></Lt>$NoTerminado</And></Where><OrderBy><FieldRef Name='FechaLimite'/></OrderBy>" }
    @{ Titulo = 'Mis trámites'
       Campos = @('LinkTitle', 'NumExpediente', 'Cliente', 'FechaLimite', 'EstadoTramite', 'Prioridad')
       Query  = "<Where><And><Eq><FieldRef Name='ResponsableTram'/><Value Type='Integer'><UserID Type='Integer'/></Value></Eq>$NoTerminado</And></Where><OrderBy><FieldRef Name='FechaLimite'/></OrderBy>" }
    @{ Titulo = 'Por expediente'
       Campos = @('LinkTitle', 'Cliente', 'Organismo', 'FechaLimite', 'EstadoTramite', 'FechaPresentacion', 'NumJustificante')
       Query  = "<GroupBy Collapse='TRUE'><FieldRef Name='NumExpediente'/></GroupBy><OrderBy><FieldRef Name='FechaLimite'/></OrderBy>" }
    @{ Titulo = 'Todos los trámites'
       Campos = @('LinkTitle', 'NumExpediente', 'Cliente', 'AreaTramite', 'Organismo', 'FechaLimite', 'EstadoTramite', 'ResponsableTram', 'BaseLegal')
       Query  = "<OrderBy><FieldRef Name='FechaLimite' Ascending='FALSE'/></OrderBy>" }
)

# =====================================================================================
#  FUNCIONES (no hace falta tocar nada a partir de aquí)
# =====================================================================================

function ConvertTo-XmlSeguro([string]$Texto) {
    return [System.Security.SecurityElement]::Escape($Texto)
}

function New-CampoXml([hashtable]$C) {
    $interno = $C.Interno
    $nombre  = ConvertTo-XmlSeguro $C.Nombre
    $req     = if ($C.Requerida) { 'TRUE' } else { 'FALSE' }
    $desc    = if ($C.Descripcion) { " Description=`"$(ConvertTo-XmlSeguro $C.Descripcion)`"" } else { '' }
    $comunes = "ID=`"{$([guid]::NewGuid())}`" Name=`"$interno`" StaticName=`"$interno`" DisplayName=`"$nombre`" Required=`"$req`"$desc"
    $defecto = if ($null -ne $C.Defecto) { "<Default>$(ConvertTo-XmlSeguro $C.Defecto)</Default>" } else { '' }

    switch ($C.Tipo) {
        'Texto'      { return "<Field Type=`"Text`" $comunes MaxLength=`"255`">$defecto</Field>" }
        'TextoLargo' { return "<Field Type=`"Note`" $comunes NumLines=`"6`" RichText=`"FALSE`" />" }
        'Numero'     { return "<Field Type=`"Number`" $comunes>$defecto</Field>" }
        'Moneda'     { return "<Field Type=`"Currency`" $comunes LCID=`"3082`" Decimals=`"2`">$defecto</Field>" }
        'Fecha'      { return "<Field Type=`"DateTime`" $comunes Format=`"DateOnly`" FriendlyDisplayFormat=`"Disabled`">$defecto</Field>" }
        'SiNo'       { return "<Field Type=`"Boolean`" $comunes>$defecto</Field>" }
        'Persona'    { return "<Field Type=`"User`" $comunes UserSelectionMode=`"PeopleOnly`" UserSelectionScope=`"0`" ShowField=`"ImnName`" />" }
        { $_ -in 'Eleccion', 'EleccionMultiple' } {
            $tipoSp  = if ($C.Tipo -eq 'EleccionMultiple') { 'MultiChoice' } else { 'Choice' }
            $choices = ($C.Opciones | ForEach-Object { "<CHOICE>$(ConvertTo-XmlSeguro $_)</CHOICE>" }) -join ''
            return "<Field Type=`"$tipoSp`" $comunes Format=`"Dropdown`" FillInChoice=`"FALSE`">$defecto<CHOICES>$choices</CHOICES></Field>"
        }
        default { throw "Tipo de columna desconocido: '$($C.Tipo)' en '$interno'" }
    }
}

function Add-Columnas {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$Lista, [array]$Columnas)
    $existentes = Get-PnPField -List $Lista | Select-Object -ExpandProperty InternalName
    foreach ($c in $Columnas) {
        if ($existentes -contains $c.Interno) {
            Write-Host "   = Columna '$($c.Nombre)' ($($c.Interno)) ya existe; se omite." -ForegroundColor DarkGray
            continue
        }
        if ($PSCmdlet.ShouldProcess("$Lista", "Crear columna '$($c.Nombre)' [$($c.Tipo)]")) {
            $xml = New-CampoXml $c
            $null = Add-PnPFieldFromXml -List $Lista -FieldXml $xml
            if ($c.Indexada) {
                # Índice: necesario para filtrar/ordenar listas de más de 5.000 elementos
                $null = Set-PnPField -List $Lista -Identity $c.Interno -Values @{ Indexed = $true }
            }
            Write-Host "   + Columna '$($c.Nombre)' creada." -ForegroundColor Green
        }
    }
}

function Add-Vistas {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$Lista, [array]$Vistas)
    $existentes = Get-PnPView -List $Lista | Select-Object -ExpandProperty Title
    foreach ($v in $Vistas) {
        if ($existentes -contains $v.Titulo) {
            Write-Host "   = Vista '$($v.Titulo)' ya existe; se omite." -ForegroundColor DarkGray
            continue
        }
        if ($PSCmdlet.ShouldProcess("$Lista", "Crear vista '$($v.Titulo)'")) {
            $p = @{ List = $Lista; Title = $v.Titulo; Fields = $v.Campos; Query = $v.Query; RowLimit = 100; Paged = $true }
            if ($v.Defecto) { $p.SetAsDefault = $true }
            $null = Add-PnPView @p
            Write-Host "   + Vista '$($v.Titulo)' creada." -ForegroundColor Green
        }
    }
}

# =====================================================================================
#  EJECUCIÓN
# =====================================================================================

$log = Join-Path $PSScriptRoot ("Crear-ListasExpedientes_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $log -WhatIf:$false | Out-Null

try {
    Write-Host "Conectando a $SiteUrl ..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

    # ---- 1. Biblioteca Expedientes ------------------------------------------------
    Write-Host "`n[1/2] Biblioteca '$NombreBiblioteca'" -ForegroundColor Cyan
    if (Get-PnPList -Identity $NombreBiblioteca -ErrorAction SilentlyContinue) {
        Write-Host "   = Ya existe; se revisan columnas y vistas." -ForegroundColor DarkGray
    }
    elseif ($PSCmdlet.ShouldProcess($SiteUrl, "Crear biblioteca '$NombreBiblioteca'")) {
        $null = New-PnPList -Title $NombreBiblioteca -Url $UrlBiblioteca -Template DocumentLibrary -OnQuickLaunch
        # Control de versiones: conserva historial de cada documento (trazabilidad, art. 32 RGPD)
        $null = Set-PnPList -Identity $NombreBiblioteca -EnableVersioning $true -MajorVersions 100 -EnableFolderCreation $true
        Write-Host "   + Biblioteca creada (versiones activadas, 100 versiones)." -ForegroundColor Green
    }
    if (Get-PnPList -Identity $NombreBiblioteca -ErrorAction SilentlyContinue) {
        Add-Columnas -Lista $NombreBiblioteca -Columnas $ColumnasExpedientes
        Add-Vistas   -Lista $NombreBiblioteca -Vistas   $VistasExpedientes
    }

    # ---- 2. Lista Control de Trámites ---------------------------------------------
    Write-Host "`n[2/2] Lista '$NombreLista'" -ForegroundColor Cyan
    if (Get-PnPList -Identity $NombreLista -ErrorAction SilentlyContinue) {
        Write-Host "   = Ya existe; se revisan columnas y vistas." -ForegroundColor DarkGray
    }
    elseif ($PSCmdlet.ShouldProcess($SiteUrl, "Crear lista '$NombreLista'")) {
        $null = New-PnPList -Title $NombreLista -Url $UrlLista -Template GenericList -OnQuickLaunch
        $null = Set-PnPList -Identity $NombreLista -EnableVersioning $true -MajorVersions 50
        $null = Set-PnPField -List $NombreLista -Identity 'Title' -Values @{ Title = 'Trámite' }
        Write-Host "   + Lista creada (columna Título renombrada a 'Trámite')." -ForegroundColor Green
    }
    if (Get-PnPList -Identity $NombreLista -ErrorAction SilentlyContinue) {
        Add-Columnas -Lista $NombreLista -Columnas $ColumnasTramites
        Add-Vistas   -Lista $NombreLista -Vistas   $VistasTramites
    }

    Write-Host "`nTerminado. Registro guardado en: $log" -ForegroundColor Cyan
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Nada de lo ya creado se ha borrado. Corrija el error y vuelva a ejecutar el script: continuará donde se quedó." -ForegroundColor Yellow
    throw
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}
