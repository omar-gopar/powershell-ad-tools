<#
.SYNOPSIS
    Actualización masiva de cargos y OUs de usuarios en Active Directory desde Excel.

.DESCRIPTION
    Lee un Excel proporcionado por RH con: # Empleado, Nombre Completo,
    Unidad Administrativa (OU destino) y Cargo.

    Por cada fila:
      - Busca el usuario en AD por Nombre Completo (DisplayName)
      - Actualiza su campo Title (Cargo)
      - Lo mueve a la OU de la Unidad Administrativa indicada
      - NO crea ni elimina usuarios
      - Normaliza acentos y caracteres especiales para evitar errores de búsqueda

.PARAMETER ExcelEntrada
    Ruta al Excel proporcionado por RH.

.PARAMETER ExcelSalida
    Ruta del Excel de resultados que se generará.

.PARAMETER LogPath
    Carpeta donde se guardará el log. Por defecto: directorio del script.

.PARAMETER Simular
    Si se activa, no se realizan cambios en AD (dry-run).

.EXAMPLE
    .\Actualizar-UsuariosAD.ps1 -ExcelEntrada ".\rh_cargos.xlsx" -ExcelSalida ".\resultados.xlsx"

.EXAMPLE
    .\Actualizar-UsuariosAD.ps1 -ExcelEntrada ".\rh_cargos.xlsx" -ExcelSalida ".\resultados.xlsx" -Simular
#>

[CmdletBinding()]
param (
    [string]$ExcelEntrada,
    [string]$ExcelSalida,
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),
    [switch]$Simular
)

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN GLOBAL
# ═══════════════════════════════════════════════════════════════════════════════
$ErrorActionPreference = "Stop"
$global:ModoSim = $Simular.IsPresent
$TS             = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:LogFile = Join-Path $LogPath "AD_Actualizacion_$TS.log"
$Script:Total   = 0
$Script:OK      = 0
$Script:Fail    = 0
$Script:Omit    = 0

# Mapa de columnas Excel → campo interno
# ⚠️ Si RH cambia los encabezados, solo edita los valores aquí
$COL = @{
    NumEmpleado  = '# Empleado'
    Nombre       = 'Nombre Completo'
    UnidadAdmin  = 'Unidad Administrativa'
    Cargo        = 'Cargo'
}

# Columnas obligatorias (claves internas)
$OBLIGATORIAS = @("Nombre", "UnidadAdmin", "Cargo")

# ═══════════════════════════════════════════════════════════════════════════════
#  FUNCIONES
# ═══════════════════════════════════════════════════════════════════════════════

function Write-Log {
    param(
        [string]$Msg,
        [ValidateSet("INFO","OK","WARN","ERR","SIM")][string]$L = "INFO"
    )
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $linea = "[$ts] [$L] $Msg"
    switch ($L) {
        "OK"   { Write-Host $linea -ForegroundColor Green  }
        "WARN" { Write-Host $linea -ForegroundColor Yellow }
        "ERR"  { Write-Host $linea -ForegroundColor Red    }
        "SIM"  { Write-Host $linea -ForegroundColor Cyan   }
        default{ Write-Host $linea -ForegroundColor White  }
    }
    Add-Content -Path $Script:LogFile -Value $linea -Encoding UTF8
}

function Remove-Acentos {
    # Normaliza una cadena eliminando tildes y caracteres diacríticos
    # Ej: "Dirección de Administración" → "Direccion de Administracion"
    param([string]$Texto)
    if ([string]::IsNullOrWhiteSpace($Texto)) { return "" }
    $normalizado = $Texto.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = [System.Text.StringBuilder]::new()
    foreach ($c in $normalizado.ToCharArray()) {
        $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

function Get-Val {
    param([hashtable]$Row, [string]$Col)
    if (-not $Row.ContainsKey($Col) -or $null -eq $Row[$Col]) { return "" }
    return "$($Row[$Col])".Trim()
}

function Find-ADUsuario {
    # Busca un usuario en AD por DisplayName, con y sin acentos
    param([string]$NombreCompleto)

    # Intento 1: búsqueda directa por DisplayName
    $resultado = Get-ADUser -Filter "DisplayName -eq '$NombreCompleto'" `
                     -Properties DisplayName, Title, Department, DistinguishedName `
                     -ErrorAction SilentlyContinue |
                 Select-Object -First 1
    if ($resultado) { return $resultado }

    # Intento 2: nombre sin acentos (por si AD tiene el nombre sin tildes)
    $sinAcento = Remove-Acentos $NombreCompleto
    if ($sinAcento -ne $NombreCompleto) {
        $resultado = Get-ADUser -Filter "DisplayName -eq '$sinAcento'" `
                         -Properties DisplayName, Title, Department, DistinguishedName `
                         -ErrorAction SilentlyContinue |
                     Select-Object -First 1
        if ($resultado) { return $resultado }
    }

    # Intento 3: buscar por Name (que en AD puede diferir del DisplayName)
    $resultado = Get-ADUser -Filter "Name -eq '$NombreCompleto'" `
                     -Properties DisplayName, Title, Department, DistinguishedName `
                     -ErrorAction SilentlyContinue |
                 Select-Object -First 1
    if ($resultado) { return $resultado }

    # Intento 4: Name sin acentos
    if ($sinAcento -ne $NombreCompleto) {
        $resultado = Get-ADUser -Filter "Name -eq '$sinAcento'" `
                         -Properties DisplayName, Title, Department, DistinguishedName `
                         -ErrorAction SilentlyContinue |
                     Select-Object -First 1
        if ($resultado) { return $resultado }
    }

    return $null
}

function Get-OuDestino {
    # Busca la OU en AD por nombre, con y sin acentos
    param([string]$NombreOU, [string]$BaseDN)

    $variantes = @(
        $NombreOU
        (Remove-Acentos $NombreOU)
    ) | Select-Object -Unique

    foreach ($v in $variantes) {
        $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$v'" `
                  -SearchBase $BaseDN -ErrorAction SilentlyContinue |
              Select-Object -First 1
        if ($ou) {
            Write-Log "   🗂️  OU encontrada: $($ou.DistinguishedName)" -L INFO
            return $ou.DistinguishedName
        }
    }

    # Si no existe, crearla (sin acentos para evitar problemas futuros)
    $nombreSeguro = Remove-Acentos $NombreOU
    if ($global:ModoSim) {
        $dnSim = "OU=$nombreSeguro,$BaseDN"
        Write-Log "   🔵 [SIM] OU '$nombreSeguro' sería creada en: $dnSim" -L SIM
        return $dnSim
    }

    try {
        $nueva = New-ADOrganizationalUnit -Name $nombreSeguro -Path $BaseDN `
                     -ProtectedFromAccidentalDeletion $false -PassThru -ErrorAction Stop
        Write-Log "   ✅ OU '$nombreSeguro' creada: $($nueva.DistinguishedName)" -L OK
        return $nueva.DistinguishedName
    } catch {
        throw "No se pudo crear la OU '$nombreSeguro': $($_.Exception.Message)"
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  INICIO — BANNER Y MODO
# ═══════════════════════════════════════════════════════════════════════════════

if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor DarkYellow
Write-Host "  ║   🔄  ACTUALIZACIÓN MASIVA DE USUARIOS — AD     ║" -ForegroundColor DarkYellow
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor DarkYellow
Write-Host ""

if ($global:ModoSim) {
    Write-Host "  🔵 MODO SIMULACIÓN ACTIVO — No se harán cambios en AD" -ForegroundColor Cyan
    Write-Log "MODO SIMULACIÓN ACTIVO" -L SIM
} else {
    Write-Host "  🟡 MODO REAL — Los cambios se aplicarán en AD" -ForegroundColor Yellow
    Write-Log "MODO REAL ACTIVO" -L WARN
}
Write-Host ""

# ── Solicitar parámetros faltantes ────────────────────────────────────────────
if (-not $ExcelEntrada) { $ExcelEntrada = Read-Host "  📂 Ruta del Excel de RH (entrada)" }
if (-not (Test-Path $ExcelEntrada -PathType Leaf)) {
    Write-Host "  ❌ Archivo no encontrado: $ExcelEntrada" -ForegroundColor Red
    exit 1
}
if (-not $ExcelSalida) { $ExcelSalida = Read-Host "  💾 Ruta del Excel de resultados (salida)" }

Write-Log "Excel entrada : $ExcelEntrada" -L INFO
Write-Log "Excel salida  : $ExcelSalida"  -L INFO
Write-Log "Log           : $Script:LogFile" -L INFO

# ═══════════════════════════════════════════════════════════════════════════════
#  MÓDULOS
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "  🔧 Verificando módulos..." -ForegroundColor White

try {
    Import-Module ActiveDirectory -WarningAction SilentlyContinue -ErrorAction Stop
    Write-Host "  ✅ ActiveDirectory cargado" -ForegroundColor Green
    Write-Log "Módulo ActiveDirectory OK" -L OK
} catch {
    Write-Host "  ❌ Módulo ActiveDirectory no disponible." -ForegroundColor Red
    Write-Host "     Ejecuta: Install-WindowsFeature RSAT-AD-PowerShell" -ForegroundColor Yellow
    Write-Log "ActiveDirectory no disponible: $_" -L ERR
    exit 1
}

try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "  📦 Instalando ImportExcel..." -ForegroundColor Yellow
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -WarningAction SilentlyContinue -ErrorAction Stop
    Write-Host "  ✅ ImportExcel cargado" -ForegroundColor Green
    Write-Log "Módulo ImportExcel OK" -L OK
} catch {
    Write-Host "  ❌ ImportExcel no disponible: $_" -ForegroundColor Red
    Write-Log "ImportExcel no disponible: $_" -L ERR
    exit 1
}

# ═══════════════════════════════════════════════════════════════════════════════
#  AUTODETECCIÓN DEL DOMINIO
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  🌐 Detectando dominio de AD..." -ForegroundColor White

try {
    $adDomain = Get-ADDomain -ErrorAction Stop
    $DomainDN = $adDomain.DistinguishedName
    Write-Host "  ✅ Dominio: $($adDomain.DNSRoot) ($DomainDN)" -ForegroundColor Green
    Write-Log "Dominio AD: $($adDomain.DNSRoot) | DN: $DomainDN" -L OK
} catch {
    Write-Host "  ⚠️  No se pudo autodetectar el dominio." -ForegroundColor Yellow
    $DomainDN = Read-Host "  🏢 DN raíz del dominio (ej: DC=INFO-DF,DC=ORG,DC=MX)"
    Write-Log "DN ingresado manualmente: $DomainDN" -L WARN
}

# ═══════════════════════════════════════════════════════════════════════════════
#  LECTURA DEL EXCEL (EPPlus directo)
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  📂 Leyendo Excel de RH..." -ForegroundColor White
Write-Log "Leyendo: $ExcelEntrada" -L INFO

$Registros = @()

try {
    $ruta = (Resolve-Path $ExcelEntrada).Path
    $pck  = New-Object OfficeOpenXml.ExcelPackage -ArgumentList (New-Object System.IO.FileInfo $ruta)
    $ws   = $pck.Workbook.Worksheets[1]

    if ($null -eq $ws -or $null -eq $ws.Dimension) { throw "La hoja está vacía o no existe." }

    $totalRows = $ws.Dimension.Rows
    $totalCols = $ws.Dimension.Columns

    Write-Host "  📊 Hoja: '$($ws.Name)' | Filas: $totalRows | Columnas: $totalCols" -ForegroundColor White

    if ($totalRows -lt 2) { throw "El Excel no tiene filas de datos (solo encabezado)." }

    # Leer encabezados fila 1
    $headers = @{}
    for ($c = 1; $c -le $totalCols; $c++) {
        $val = $ws.Cells[1, $c].Text
        if ($val -and $val.Trim() -ne "") { $headers[$c] = $val.Trim() }
    }

    Write-Host "  🏷️  Columnas: $($headers.Values -join ' | ')" -ForegroundColor White
    Write-Log "Columnas detectadas: $($headers.Values -join ' | ')" -L INFO

    # --- MEJORA: Validar que los nombres de columna esperados existan ---
    # Primero, aseguramos que $COL no tenga valores vacíos (fallback)
    $expectedColumns = @()
    foreach ($campo in $OBLIGATORIAS) {
        $colName = $COL[$campo]
        if ([string]::IsNullOrWhiteSpace($colName)) {
            # Si está vacío, asumimos el nombre esperado según el campo
            switch ($campo) {
                "Nombre"      { $colName = "Nombre Completo" }
                "UnidadAdmin" { $colName = "Unidad Administrativa" }
                "Cargo"       { $colName = "Cargo" }
                default       { $colName = $campo }
            }
            Write-Log "⚠️  COL[$campo] estaba vacío. Se usará '$colName'" -L WARN
            $COL[$campo] = $colName   # actualizamos el mapa para el resto del script
        }
        $expectedColumns += $colName
    }

    $presentes = $headers.Values
    foreach ($req in $expectedColumns) {
        if ($req -notin $presentes) {
            throw "Columna obligatoria no encontrada: '$req'. Presentes: $($presentes -join ', ')"
        }
    }

    # Índice inverso: nombre → número de columna
    $colIdx = @{}
    foreach ($c in $headers.Keys) { $colIdx[$headers[$c]] = $c }

    # Leer filas de datos
    for ($r = 2; $r -le $totalRows; $r++) {
        $row = @{}
        foreach ($c in $headers.Keys) { $row[$headers[$c]] = $ws.Cells[$r, $c].Text }

        $chkNombre = if ($colIdx.ContainsKey($COL["Nombre"])) { $row[$COL["Nombre"]] } else { "" }
        if (-not $chkNombre) { continue }   # omitir filas vacías

        $Registros += [PSCustomObject]@{
            NumEmpleado = Get-Val $row $COL["NumEmpleado"]
            Nombre      = Get-Val $row $COL["Nombre"]
            UnidadAdmin = Get-Val $row $COL["UnidadAdmin"]
            Cargo       = Get-Val $row $COL["Cargo"]
            _Fila       = $r
        }
    }

    $pck.Dispose()
    Write-Host "  ✅ Registros leídos: $($Registros.Count)" -ForegroundColor Green
    Write-Log "Registros leídos: $($Registros.Count)" -L OK

} catch {
    Write-Host "  ❌ Error al leer el Excel: $_" -ForegroundColor Red
    Write-Log "Error al leer el Excel: $_" -L ERR
    exit 1
}

if ($Registros.Count -eq 0) {
    Write-Host "  ❌ No hay registros para procesar." -ForegroundColor Red
    Write-Log "Sin registros para procesar." -L ERR
    exit 1
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PROCESO PRINCIPAL — ACTUALIZACIÓN
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  ⚙️  PROCESANDO ACTUALIZACIONES..." -ForegroundColor White
Write-Log "═══ INICIO DE ACTUALIZACIONES ═══" -L INFO

$Resultados = @()
$Errores    = @()

foreach ($R in $Registros) {
    $Script:Total++
    $empID = if ($R.NumEmpleado) { "#$($R.NumEmpleado) " } else { "" }

    Write-Host ""
    Write-Host "  👤 Fila $($R._Fila): $empID$($R.Nombre)" -ForegroundColor White
    Write-Log "── Fila $($R._Fila): $empID$($R.Nombre) ──" -L INFO
	
    # Validar campos obligatorios
    $errVal = @()
    foreach ($campo in $OBLIGATORIAS) {
        if ([string]::IsNullOrWhiteSpace($R.$campo)) {
            $errVal += "Campo vacío: $campo"
        }
    }
    if ($errVal.Count -gt 0) {
        $msgErr = $errVal -join " | "
        Write-Host "     ⛔ Validación fallida: $msgErr" -ForegroundColor Red
        Write-Log "VALIDACION FALLIDA | Fila $($R._Fila) | $msgErr" -L ERR
        $Script:Fail++
        $obj = [PSCustomObject]@{
            Fila        = $R._Fila
            NumEmpleado = $R.NumEmpleado
            Nombre      = $R.Nombre
            UnidadAdmin = $R.UnidadAdmin
            Cargo       = $R.Cargo
            SamAccount  = ""
            OUAnterior  = ""
            OUNueva     = ""
            CargoAnterior = ""
            Estado      = "ERROR_VALIDACION"
            Error       = $msgErr
            FechaHora   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $Resultados += $obj
        $Errores    += $obj
        continue
    }

    try {
        # ── 1. Buscar usuario en AD ──────────────────────────────────────────
        $adUser = Find-ADUsuario -NombreCompleto $R.Nombre

        if ($null -eq $adUser) {
            throw "Usuario no encontrado en AD con DisplayName/Name '$($R.Nombre)' (con y sin acentos)"
        }

        Write-Host "     🔍 Encontrado: $($adUser.SamAccountName) | DN: $($adUser.DistinguishedName)" -ForegroundColor White
        Write-Log "   Encontrado: $($adUser.SamAccountName) | DN actual: $($adUser.DistinguishedName)" -L INFO

        $ouActualDN    = ($adUser.DistinguishedName -split ',', 2)[1]
        $cargoActual   = $adUser.Title

        # ── 2. Resolver OU destino ───────────────────────────────────────────
        $ouNuevaDN = Get-OuDestino -NombreOU $R.UnidadAdmin -BaseDN $DomainDN

        # ── 3. Detectar si hay cambios reales ────────────────────────────────
        $cambiaOU    = ($ouActualDN -ne $ouNuevaDN)
        $cambiaCargo = ($cargoActual -ne $R.Cargo)
		$empNumberActual = $adUser.EmployeeNumber
		$cambiaEmpNum    = ($empNumberActual -ne $R.NumEmpleado)

        if (-not $cambiaOU -and -not $cambiaCargo -and -not $cambiaEmpNum) {
            Write-Host "     ⏭️  Sin cambios necesarios (cargo y OU ya están actualizados)" -ForegroundColor DarkGray
            Write-Log "SIN CAMBIOS | $($adUser.SamAccountName) | Cargo y OU ya correctos" -L INFO
            $Script:Omit++
            $Resultados += [PSCustomObject]@{
                Fila          = $R._Fila
                NumEmpleado   = $R.NumEmpleado
                Nombre        = $R.Nombre
                UnidadAdmin   = $R.UnidadAdmin
                Cargo         = $R.Cargo
                SamAccount    = $adUser.SamAccountName
                OUAnterior    = $ouActualDN
                OUNueva       = $ouNuevaDN
                CargoAnterior = $cargoActual
                Estado        = "SIN_CAMBIOS"
                Error         = ""
                FechaHora     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            continue
        }

        # ── 4. Aplicar cambios ───────────────────────────────────────────────
        if ($global:ModoSim) {
            if ($cambiaCargo) {
                Write-Host "     🔵 [SIM] Cargo: '$cargoActual' → '$($R.Cargo)'" -ForegroundColor Cyan
                Write-Log "SIM | Cargo: '$cargoActual' → '$($R.Cargo)'" -L SIM
            }
            if ($cambiaOU) {
                Write-Host "     🔵 [SIM] Mover: '$ouActualDN' → '$ouNuevaDN'" -ForegroundColor Cyan
                Write-Log "SIM | Mover: '$ouActualDN' → '$ouNuevaDN'" -L SIM
            }
            $estado = "SIMULADO"
			
			if ($cambiaEmpNum) {
			Write-Host "     🔵 [SIM] EmployeeNumber: '$empNumberActual' → '$($R.NumEmpleado)'" -ForegroundColor Cyan
			Write-Log "SIM | EmployeeNumber: '$empNumberActual' → '$($R.NumEmpleado)'" -L SIM
			}
			$estado = "SIMULADO"

        } else {
            # Actualizar cargo (Title)
            if ($cambiaCargo) {
                Set-ADUser -Identity $adUser.SamAccountName -Title $R.Cargo -ErrorAction Stop
                Write-Host "     ✅ Cargo actualizado: '$cargoActual' → '$($R.Cargo)'" -ForegroundColor Green
                Write-Log "Cargo actualizado: '$cargoActual' → '$($R.Cargo)' | $($adUser.SamAccountName)" -L OK
            }

            # Mover a nueva OU
            if ($cambiaOU) {
                Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $ouNuevaDN -ErrorAction Stop
                Write-Host "     ✅ Movido a OU: $ouNuevaDN" -ForegroundColor Green
                Write-Log "Movido: '$ouActualDN' → '$ouNuevaDN' | $($adUser.SamAccountName)" -L OK
            }
			
			if ($cambiaEmpNum) {
			Set-ADUser -Identity $adUser.SamAccountName -EmployeeNumber $R.NumEmpleado -ErrorAction Stop
			Write-Host "     ✅ EmployeeNumber actualizado: '$empNumberActual' → '$($R.NumEmpleado)'" -ForegroundColor Green
			Write-Log "EmployeeNumber actualizado: '$empNumberActual' → '$($R.NumEmpleado)' | $($adUser.SamAccountName)" -L OK
			}

            $estado = "ACTUALIZADO"
        }

        $Script:OK++
        $Resultados += [PSCustomObject]@{
            Fila          = $R._Fila
            NumEmpleado   = $R.NumEmpleado
            Nombre        = $R.Nombre
            UnidadAdmin   = $R.UnidadAdmin
            Cargo         = $R.Cargo
            SamAccount    = $adUser.SamAccountName
            OUAnterior    = $ouActualDN
            OUNueva       = $ouNuevaDN
            CargoAnterior = $cargoActual
            Estado        = $estado
            Error         = ""
            FechaHora     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }

    } catch {
        $msgErr = $_.Exception.Message
        Write-Host "     ❌ Error: $msgErr" -ForegroundColor Red
        Write-Log "ERROR | Fila $($R._Fila) | $($R.Nombre) | $msgErr" -L ERR
        $Script:Fail++
        $obj = [PSCustomObject]@{
            Fila          = $R._Fila
            NumEmpleado   = $R.NumEmpleado
            Nombre        = $R.Nombre
            UnidadAdmin   = $R.UnidadAdmin
            Cargo         = $R.Cargo
            SamAccount    = ""
            OUAnterior    = ""
            OUNueva       = ""
            CargoAnterior = ""
            Estado        = "ERROR"
            Error         = $msgErr
            FechaHora     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $Resultados += $obj
        $Errores    += $obj
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORTAR RESULTADOS A EXCEL (3 hojas)
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  💾 Exportando resultados..." -ForegroundColor White

try {
    $actualizados = $Resultados | Where-Object { $_.Estado -in @("ACTUALIZADO","SIMULADO") }
    $sinCambios   = $Resultados | Where-Object { $_.Estado -eq "SIN_CAMBIOS" }
    $fallidos     = $Resultados | Where-Object { $_.Estado -notin @("ACTUALIZADO","SIMULADO","SIN_CAMBIOS") }

    $colsActualiz = "Fila","NumEmpleado","Nombre","SamAccount","CargoAnterior","Cargo","OUAnterior","OUNueva","Estado","FechaHora"
    $colsError    = "Fila","NumEmpleado","Nombre","UnidadAdmin","Cargo","Estado","Error","FechaHora"
    $colsSinCamb  = "Fila","NumEmpleado","Nombre","SamAccount","Cargo","OUNueva","Estado","FechaHora"

    if ($actualizados) {
        $actualizados | Select-Object $colsActualiz |
            Export-Excel -Path $ExcelSalida -WorksheetName "Actualizados" `
                         -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium6
    }
    if ($sinCambios) {
        $sinCambios | Select-Object $colsSinCamb |
            Export-Excel -Path $ExcelSalida -WorksheetName "Sin Cambios" `
                         -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium2 -Append
    }
    if ($fallidos) {
        $fallidos | Select-Object $colsError |
            Export-Excel -Path $ExcelSalida -WorksheetName "Errores" `
                         -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium3 -Append
    }

    Write-Host "  ✅ Excel exportado: $ExcelSalida" -ForegroundColor Green
    Write-Log "Excel exportado: $ExcelSalida" -L OK

} catch {
    Write-Host "  ⚠️  No se pudo exportar Excel: $_" -ForegroundColor Yellow
    $csv = $ExcelSalida -replace '\.xlsx$', '.csv'
    $Resultados | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host "  💾 Fallback CSV: $csv" -ForegroundColor Yellow
    Write-Log "Fallback CSV: $csv" -L WARN
}

# ═══════════════════════════════════════════════════════════════════════════════
#  RESUMEN FINAL
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor DarkYellow
Write-Host "  ║              📊  RESUMEN FINAL                  ║" -ForegroundColor DarkYellow
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor DarkYellow
Write-Host "     📋 Total procesados  : $Script:Total"  -ForegroundColor White
Write-Host "     ✅ Actualizados      : $Script:OK"     -ForegroundColor Green
Write-Host "     ⏭️  Sin cambios       : $Script:Omit"  -ForegroundColor DarkGray
Write-Host "     ❌ Con errores       : $Script:Fail"   -ForegroundColor $(if ($Script:Fail -gt 0) {"Red"} else {"Green"})
Write-Host "     📄 Log detallado     : $Script:LogFile"  -ForegroundColor White
Write-Host "     💾 Resultados Excel  : $ExcelSalida"   -ForegroundColor White
Write-Host ""

if ($global:ModoSim) {
    Write-Host "  🔵 Esta fue una SIMULACIÓN. No se modificó AD." -ForegroundColor Cyan
    Write-Log "FIN SIMULACIÓN — Total: $Script:Total | Sim: $Script:OK | Sin cambios: $Script:Omit | Errores: $Script:Fail" -L SIM
} else {
    Write-Host "  ⚠️  Se realizaron cambios en Active Directory." -ForegroundColor Yellow
    Write-Log "FIN — Total: $Script:Total | OK: $Script:OK | Sin cambios: $Script:Omit | Errores: $Script:Fail" -L OK
}
Write-Host ""