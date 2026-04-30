<#
.SYNOPSIS
    Creación masiva de usuarios en Active Directory desde un archivo Excel.

.DESCRIPTION
    Lee un archivo Excel y crea usuarios en AD, asignándolos a la OU
    correspondiente al campo "Departamento". Soporta modo simulación,
    genera contraseñas seguras y exporta resultados en Excel + log detallado.

    Columnas esperadas en el Excel:
      Nombre, Apellido, Usuario, Departamento, Compania,
      Calle, Ciudad, Estado, CP, Pais, Email, Telefono

.PARAMETER ExcelEntrada
    Ruta al archivo Excel con los datos de los usuarios.

.PARAMETER ExcelSalida
    Ruta del Excel de resultados que se generará.

.PARAMETER DominioUPN
    Dominio para el UPN. Ej: empresa.dominio.com

.PARAMETER LogPath
    Carpeta donde se guardará el log. Por defecto: directorio del script.

.PARAMETER Simular
    Switch. Si se activa, no se realizan cambios en AD (dry-run).

.EXAMPLE
    .\Crear-UsuariosAD.ps1 -ExcelEntrada ".\usuarios.xlsx" -ExcelSalida ".\resultados.xlsx" -DominioUPN "empresa.dominio.com"

.EXAMPLE
    .\Crear-UsuariosAD.ps1 -ExcelEntrada ".\usuarios.xlsx" -ExcelSalida ".\resultados.xlsx" -DominioUPN "empresa.dominio.com" -Simular
#>

[CmdletBinding()]
param (
    [string]$ExcelEntrada,
    [string]$ExcelSalida,
    [string]$DominioUPN,
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),
    [switch]$Simular
)

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN GLOBAL
# ═══════════════════════════════════════════════════════════════════════════════
$ErrorActionPreference = "Stop"
$global:ModoSim = $Simular.IsPresent
$TS             = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:LogFile = Join-Path $LogPath "AD_Log_$TS.log"
$Script:Total   = 0
$Script:OK      = 0
$Script:Fail    = 0

# Mapa de columnas Excel → campo interno
# ⚠️ Si cambias encabezados en el Excel, solo edita los valores aquí
$COL = @{
    Nombre       = "Nombre"
    Apellido     = "Apellido"
    Usuario      = "Usuario"
    Departamento = "Departamento"
    Compania     = "Compania"
    Calle        = "Calle"
    Ciudad       = "Ciudad"
    Estado       = "Estado"
    CP           = "CP"
    Pais         = "Pais"
    Email        = "Email"
    Telefono     = "Telefono"
}
$OBLIGATORIAS = @("Nombre","Apellido","Usuario","Departamento","Email")

# ═══════════════════════════════════════════════════════════════════════════════
#  FUNCIONES
# ═══════════════════════════════════════════════════════════════════════════════

function Write-Log {
    param(
        [string]$Msg,
        [ValidateSet("INFO","OK","WARN","ERR","SIM")][string]$L = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $linea     = "[$timestamp] [$L] $Msg"

    switch ($L) {
        "OK"   { Write-Host $linea -ForegroundColor Green  }
        "WARN" { Write-Host $linea -ForegroundColor Yellow }
        "ERR"  { Write-Host $linea -ForegroundColor Red    }
        "SIM"  { Write-Host $linea -ForegroundColor Cyan   }
        default{ Write-Host $linea -ForegroundColor White  }
    }
    Add-Content -Path $Script:LogFile -Value $linea -Encoding UTF8
}

function New-Password {
    # Usa RandomNumberGenerator (.NET Cryptography) — más seguro que Get-Random
    $chars = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%&*?+'
    $rng   = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = [byte[]]::new(14)
    $rng.GetBytes($bytes)
    # Garantizar al menos 1 de cada categoría reemplazando posiciones fijas
    $pwd = -join ($bytes | ForEach-Object { $chars[$_ % $chars.Length] })
    $rng.Dispose()
    return $pwd
}

function Get-OU {
    param([string]$Nombre, [string]$Base)

    # Buscar OU existente por nombre en todo el árbol
    $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$Nombre'" `
              -SearchBase $Base -ErrorAction SilentlyContinue |
          Select-Object -First 1

    if ($ou) {
        Write-Log "   🗂️  OU encontrada: $($ou.DistinguishedName)" -L INFO
        return $ou.DistinguishedName
    }

    if ($global:ModoSim) {
        $dnSim = "OU=$Nombre,$Base"
        Write-Log "   🔵 [SIM] OU '$Nombre' sería creada en: $dnSim" -L SIM
        return $dnSim
    }

    try {
        $nueva = New-ADOrganizationalUnit -Name $Nombre -Path $Base `
                     -ProtectedFromAccidentalDeletion $false -PassThru -ErrorAction Stop
        Write-Log "   ✅ OU '$Nombre' creada: $($nueva.DistinguishedName)" -L OK
        return $nueva.DistinguishedName
    } catch {
        throw "No se pudo crear la OU '$Nombre': $($_.Exception.Message)"
    }
}

function Get-Val {
    param([hashtable]$Row, [string]$Col)
    if (-not $Row.ContainsKey($Col) -or $null -eq $Row[$Col]) { return "" }
    return "$($Row[$Col])".Trim()
}

function Test-Usuario {
    param($U)
    $err = @()
    foreach ($c in $OBLIGATORIAS) {
        if ([string]::IsNullOrWhiteSpace($U.$c)) { $err += "Campo vacío: $c" }
    }
    if ($U.Usuario -and $U.Usuario.Length -gt 20) {
        $err += "Usuario excede 20 caracteres (SAMAccountName limit)"
    }
    if ($U.Usuario -match '[\\/:*?"<>|@\s]') {
        $err += "Usuario contiene caracteres inválidos"
    }
    if ($U.Email -and $U.Email -notmatch '^[\w\.\-\+]+@[\w\.\-]+\.\w{2,}$') {
        $err += "Email con formato inválido: $($U.Email)"
    }
    return $err
}

# ═══════════════════════════════════════════════════════════════════════════════
#  INICIO — BANNER Y MODO
# ═══════════════════════════════════════════════════════════════════════════════

if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor DarkCyan
Write-Host "  ║   👥  CREACIÓN MASIVA DE USUARIOS — AD          ║" -ForegroundColor DarkCyan
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor DarkCyan
Write-Host ""

if ($global:ModoSim) {
    Write-Host "  🔵 MODO SIMULACIÓN ACTIVO — No se harán cambios en AD" -ForegroundColor Cyan
    Write-Log "MODO SIMULACIÓN ACTIVO" -L SIM
} else {
    Write-Host "  🔴 MODO REAL — Los cambios se aplicarán en AD" -ForegroundColor Red
    Write-Log "MODO REAL ACTIVO" -L WARN
}
Write-Host ""

# ── Solicitar parámetros que falten interactivamente ─────────────────────────
if (-not $ExcelEntrada) {
    $ExcelEntrada = Read-Host "  📂 Ruta del Excel de entrada"
}
if (-not (Test-Path $ExcelEntrada -PathType Leaf)) {
    Write-Host "  ❌ No se encontró el archivo: $ExcelEntrada" -ForegroundColor Red
    exit 1
}
if (-not $ExcelSalida) {
    $ExcelSalida = Read-Host "  💾 Ruta del Excel de resultados (salida)"
}
if (-not $DominioUPN) {
    $DominioUPN = Read-Host "  🌐 Dominio UPN (ej: empresa.dominio.com)"
}

Write-Log "Excel entrada : $ExcelEntrada"  -L INFO
Write-Log "Excel salida  : $ExcelSalida"   -L INFO
Write-Log "Dominio UPN   : $DominioUPN"    -L INFO
Write-Log "Log           : $Script:LogFile" -L INFO

# ═══════════════════════════════════════════════════════════════════════════════
#  MÓDULOS
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  🔧 Verificando módulos..." -ForegroundColor White

# ActiveDirectory
try {
    Import-Module ActiveDirectory -WarningAction SilentlyContinue -ErrorAction Stop
    Write-Host "  ✅ ActiveDirectory cargado" -ForegroundColor Green
    Write-Log "Módulo ActiveDirectory OK" -L OK
} catch {
    Write-Host "  ❌ Módulo ActiveDirectory no disponible: $_" -ForegroundColor Red
    Write-Host "     Ejecuta: Install-WindowsFeature RSAT-AD-PowerShell" -ForegroundColor Yellow
    Write-Log "Módulo ActiveDirectory no disponible: $_" -L ERR
    exit 1
}

# ImportExcel
try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "  📦 Instalando ImportExcel desde PSGallery..." -ForegroundColor Yellow
        Write-Log "Instalando ImportExcel..." -L WARN
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -WarningAction SilentlyContinue -ErrorAction Stop
    Write-Host "  ✅ ImportExcel cargado" -ForegroundColor Green
    Write-Log "Módulo ImportExcel OK" -L OK
} catch {
    Write-Host "  ❌ ImportExcel no disponible: $_" -ForegroundColor Red
    Write-Host "     Instala con: Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor Yellow
    Write-Log "Módulo ImportExcel no disponible: $_" -L ERR
    exit 1
}

# ═══════════════════════════════════════════════════════════════════════════════
#  AUTODETECCIÓN DEL DOMINIO
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  🌐 Detectando dominio de AD..." -ForegroundColor White

try {
    $adDomain  = Get-ADDomain -ErrorAction Stop
    $DomainDN  = $adDomain.DistinguishedName
    $DomainDNS = $adDomain.DNSRoot
    Write-Host "  ✅ Dominio detectado: $DomainDNS ($DomainDN)" -ForegroundColor Green
    Write-Log "Dominio AD: $DomainDNS | DN: $DomainDN" -L OK

    # Validar que el DominioUPN exista en el bosque
    $upnSuffixes = @($DomainDNS) + @($adDomain.UPNSuffixes)
    if ($DominioUPN -notin $upnSuffixes) {
        Write-Host "  ⚠️  El DominioUPN '$DominioUPN' no está en los sufijos UPN registrados." -ForegroundColor Yellow
        Write-Host "     Sufijos disponibles: $($upnSuffixes -join ', ')" -ForegroundColor Yellow
        Write-Log "DominioUPN '$DominioUPN' no encontrado en sufijos UPN del bosque. Sufijos: $($upnSuffixes -join ', ')" -L WARN
    }
} catch {
    Write-Host "  ⚠️  No se pudo autodetectar el dominio: $_" -ForegroundColor Yellow
    Write-Log "No se pudo autodetectar el dominio: $_" -L WARN
    $DomainDN = Read-Host "  🏢 DN raíz del dominio (ej: DC=EMPRESA,DC=DOMINIO,DC=COM)"
    Write-Log "DN ingresado manualmente: $DomainDN" -L INFO
}

# ═══════════════════════════════════════════════════════════════════════════════
#  LECTURA DEL EXCEL (EPPlus directo — sin problemas con tildes)
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  📂 Leyendo archivo Excel..." -ForegroundColor White
Write-Log "Leyendo Excel: $ExcelEntrada" -L INFO

$Usuarios = @()

try {
    $ruta = (Resolve-Path $ExcelEntrada).Path
    $pck  = New-Object OfficeOpenXml.ExcelPackage -ArgumentList (New-Object System.IO.FileInfo $ruta)
    $ws   = $pck.Workbook.Worksheets[1]

    if ($null -eq $ws)          { throw "No se encontró la Hoja 1 en el Excel." }
    if ($null -eq $ws.Dimension){ throw "La hoja está completamente vacía." }

    $totalRows = $ws.Dimension.Rows
    $totalCols = $ws.Dimension.Columns

    Write-Host "  📊 Hoja: '$($ws.Name)' | Filas: $totalRows | Columnas: $totalCols" -ForegroundColor White
    Write-Log "Hoja: $($ws.Name) | Filas: $totalRows | Columnas: $totalCols" -L INFO

    if ($totalRows -lt 2) { throw "El Excel tiene solo $totalRows fila(s). Se necesitan encabezado + datos." }

    # Leer encabezados (fila 1)
    $headers = @{}
    for ($c = 1; $c -le $totalCols; $c++) {
        $val = $ws.Cells[1, $c].Text
        if ($val -and $val.Trim() -ne "") { $headers[$c] = $val.Trim() }
    }

    Write-Host "  🏷️  Columnas: $($headers.Values -join ' | ')" -ForegroundColor White
    Write-Log "Columnas detectadas: $($headers.Values -join ' | ')" -L INFO

    # Validar columnas obligatorias
    $presentes = $headers.Values
    foreach ($campo in $OBLIGATORIAS) {
        $col = $COL[$campo]
        if ($col -notin $presentes) {
            throw "Columna obligatoria no encontrada: '$col'. Presentes: $($presentes -join ', ')"
        }
    }

    # Índice inverso: nombre → número de columna
    $colIdx = @{}
    foreach ($c in $headers.Keys) { $colIdx[$headers[$c]] = $c }

    # Leer filas de datos
    for ($r = 2; $r -le $totalRows; $r++) {
        $row = @{}
        foreach ($c in $headers.Keys) { $row[$headers[$c]] = $ws.Cells[$r, $c].Text }

        # Omitir filas vacías
        $chkNombre  = if ($colIdx.ContainsKey($COL["Nombre"]))  { $row[$COL["Nombre"]]  } else { "" }
        $chkUsuario = if ($colIdx.ContainsKey($COL["Usuario"])) { $row[$COL["Usuario"]] } else { "" }
        $chkEmail   = if ($colIdx.ContainsKey($COL["Email"]))   { $row[$COL["Email"]]   } else { "" }
        if (-not ($chkNombre -or $chkUsuario -or $chkEmail)) { continue }

        $Usuarios += [PSCustomObject]@{
            Nombre       = Get-Val $row $COL["Nombre"]
            Apellido     = Get-Val $row $COL["Apellido"]
            Usuario      = Get-Val $row $COL["Usuario"]
            Departamento = Get-Val $row $COL["Departamento"]
            Compania     = Get-Val $row $COL["Compania"]
            Calle        = Get-Val $row $COL["Calle"]
            Ciudad       = Get-Val $row $COL["Ciudad"]
            Estado       = Get-Val $row $COL["Estado"]
            CP           = Get-Val $row $COL["CP"]
            Pais         = Get-Val $row $COL["Pais"]
            Email        = Get-Val $row $COL["Email"]
            Telefono     = Get-Val $row $COL["Telefono"]
            _Fila        = $r
        }
    }

    $pck.Dispose()

    Write-Host "  ✅ Usuarios leídos: $($Usuarios.Count)" -ForegroundColor Green
    Write-Log "Usuarios leídos del Excel: $($Usuarios.Count)" -L OK

} catch {
    Write-Host "  ❌ Error al leer el Excel: $_" -ForegroundColor Red
    Write-Log "Error al leer el Excel: $_" -L ERR
    exit 1
}

if ($Usuarios.Count -eq 0) {
    Write-Host "  ❌ No hay usuarios para procesar." -ForegroundColor Red
    Write-Log "No hay usuarios para procesar." -L ERR
    exit 1
}

# ── Informe previo ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  📋 VALIDACIÓN PREVIA" -ForegroundColor White
$validosCount = 0
$invalidosCount = 0
foreach ($U in $Usuarios) {
    $errPrev = Test-Usuario -U $U
    if ($errPrev.Count -eq 0) { $validosCount++ } else { $invalidosCount++ }
}
Write-Host "     Total en Excel : $($Usuarios.Count)"  -ForegroundColor White
Write-Host "     Válidos        : $validosCount"        -ForegroundColor Green
Write-Host "     Con problemas  : $invalidosCount"      -ForegroundColor $(if ($invalidosCount -gt 0) {"Red"} else {"Green"})
Write-Log "Validación previa — Total: $($Usuarios.Count) | Válidos: $validosCount | Inválidos: $invalidosCount" -L INFO

# ═══════════════════════════════════════════════════════════════════════════════
#  PROCESO PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  ⚙️  PROCESANDO USUARIOS..." -ForegroundColor White
Write-Log "═══ INICIO DE PROCESAMIENTO ═══" -L INFO

$Resultados = @()
$Errores    = @()

foreach ($U in $Usuarios) {
    $Script:Total++
    $display = "$($U.Nombre) $($U.Apellido) [$($U.Usuario)]"
    Write-Host ""
    Write-Host "  👤 Fila $($U._Fila): $display" -ForegroundColor White
    Write-Log "── Fila $($U._Fila): $display ──" -L INFO

    # Validar
    $errVal = Test-Usuario -U $U
    if ($errVal.Count -gt 0) {
        $msgErr = $errVal -join " | "
        Write-Host "     ⛔ Validación fallida: $msgErr" -ForegroundColor Red
        Write-Log "VALIDACION FALLIDA | Fila $($U._Fila) | $msgErr" -L ERR
        $Script:Fail++
        $obj = [PSCustomObject]@{
            Fila           = $U._Fila
            Usuario        = $U.Usuario
            NombreCompleto = "$($U.Nombre) $($U.Apellido)"
            Email          = $U.Email
            Departamento   = $U.Departamento
            OU             = ""
            Password       = ""
            Telefono       = $U.Telefono
            Estado         = "ERROR_VALIDACION"
            Error          = $msgErr
            FechaHora      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $Resultados += $obj
        $Errores    += $obj
        continue
    }

    # Procesar
    try {
        $sam  = $U.Usuario.ToLower()
        $upn  = "$sam@$DominioUPN"
        $name = "$($U.Nombre) $($U.Apellido)".Trim()

        # Verificar duplicados (con -Filter, no con excepciones)
        if (Get-ADUser -Filter "SamAccountName -eq '$sam'" -EA SilentlyContinue) {
            throw "La cuenta '$sam' ya existe en AD"
        }
        if (Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -EA SilentlyContinue) {
            throw "El UPN '$upn' ya está en uso"
        }

        $ou   = Get-OU -Nombre $U.Departamento -Base $DomainDN
        $pass = New-Password
        $sec  = ConvertTo-SecureString $pass -AsPlainText -Force

        $params = @{
            Name                  = $name
            GivenName             = $U.Nombre
            Surname               = $U.Apellido
            SamAccountName        = $sam
            UserPrincipalName     = $upn
            EmailAddress          = $U.Email
            DisplayName           = $name
            AccountPassword       = $sec
            Path                  = $ou
            Enabled               = $true
            ChangePasswordAtLogon = $true
            PasswordNeverExpires  = $false
        }
        if ($U.Departamento) { $params["Department"]    = $U.Departamento }
        if ($U.Compania)     { $params["Company"]        = $U.Compania     }
        if ($U.Ciudad)       { $params["City"]           = $U.Ciudad       }
        if ($U.Estado)       { $params["State"]          = $U.Estado       }
        if ($U.CP)           { $params["PostalCode"]     = $U.CP           }
        if ($U.Pais)         { $params["Country"]        = $U.Pais         }
        if ($U.Calle)        { $params["StreetAddress"]  = $U.Calle        }
        if ($U.Telefono)     { $params["OfficePhone"]    = $U.Telefono     }

        if ($global:ModoSim) {
            Write-Host "     🔵 [SIM] Se crearía: $name | UPN: $upn | OU: $ou" -ForegroundColor Cyan
            Write-Log "SIM | $name | UPN: $upn | OU: $ou" -L SIM
            $estado = "SIMULADO"
        } else {
            New-ADUser @params -ErrorAction Stop
            Write-Host "     ✅ Usuario creado: $sam | OU: $ou" -ForegroundColor Green
            Write-Log "CREADO | $sam | UPN: $upn | OU: $ou" -L OK
            $estado = "CREADO"
        }

        $Script:OK++
        $Resultados += [PSCustomObject]@{
            Fila           = $U._Fila
            Usuario        = $sam
            NombreCompleto = $name
            Email          = $U.Email
            Departamento   = $U.Departamento
            OU             = $ou
            Password       = $pass
            Telefono       = $U.Telefono
            Estado         = $estado
            Error          = ""
            FechaHora      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }

    } catch {
        $msgErr = $_.Exception.Message
        Write-Host "     ❌ Error: $msgErr" -ForegroundColor Red
        Write-Log "ERROR | Fila $($U._Fila) | $($U.Usuario) | $msgErr" -L ERR
        $Script:Fail++
        $obj = [PSCustomObject]@{
            Fila           = $U._Fila
            Usuario        = $U.Usuario
            NombreCompleto = "$($U.Nombre) $($U.Apellido)"
            Email          = $U.Email
            Departamento   = $U.Departamento
            OU             = ""
            Password       = ""
            Telefono       = $U.Telefono
            Estado         = "ERROR_CREACION"
            Error          = $msgErr
            FechaHora      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $Resultados += $obj
        $Errores    += $obj
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORTAR RESULTADOS A EXCEL (2 hojas: Usuarios + Errores)
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  💾 Exportando resultados..." -ForegroundColor White

try {
    $exitosos = $Resultados | Where-Object { $_.Estado -in @("CREADO","SIMULADO") }
    $fallidos = $Resultados | Where-Object { $_.Estado -notIn @("CREADO","SIMULADO") }

    if ($exitosos) {
        $exitosos | Select-Object Fila,Usuario,NombreCompleto,Email,Departamento,OU,Password,Telefono,Estado,FechaHora |
            Export-Excel -Path $ExcelSalida -WorksheetName "Usuarios Creados" `
                         -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium6
    }
    if ($fallidos) {
        $fallidos | Select-Object Fila,Usuario,NombreCompleto,Email,Departamento,Estado,Error,FechaHora |
            Export-Excel -Path $ExcelSalida -WorksheetName "Errores" `
                         -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium3 -Append
    }

    Write-Host "  ✅ Excel exportado: $ExcelSalida" -ForegroundColor Green
    Write-Log "Excel de resultados exportado: $ExcelSalida" -L OK
} catch {
    Write-Host "  ⚠️  No se pudo exportar el Excel: $_" -ForegroundColor Yellow
    Write-Log "Error exportando Excel: $_" -L WARN
    # Fallback a CSV
    $fallbackCsv = $ExcelSalida -replace '\.xlsx$', '.csv'
    $Resultados | Export-Csv -Path $fallbackCsv -NoTypeInformation -Encoding UTF8
    Write-Host "  💾 Fallback: resultados guardados como CSV en $fallbackCsv" -ForegroundColor Yellow
    Write-Log "Fallback CSV: $fallbackCsv" -L WARN
}

# ═══════════════════════════════════════════════════════════════════════════════
#  RESUMEN FINAL
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════╗" -ForegroundColor DarkCyan
Write-Host "  ║              📊  RESUMEN FINAL                  ║" -ForegroundColor DarkCyan
Write-Host "  ╚══════════════════════════════════════════════════╝" -ForegroundColor DarkCyan
Write-Host "     👥 Total procesados  : $Script:Total"  -ForegroundColor White
Write-Host "     ✅ Exitosos          : $Script:OK"     -ForegroundColor Green
Write-Host "     ❌ Con errores       : $Script:Fail"   -ForegroundColor $(if ($Script:Fail -gt 0) {"Red"} else {"Green"})
Write-Host "     📄 Log detallado     : $Script:LogFile"  -ForegroundColor White
Write-Host "     💾 Resultados Excel  : $ExcelSalida"   -ForegroundColor White
Write-Host ""

if ($global:ModoSim) {
    Write-Host "  🔵 Esta fue una SIMULACIÓN. No se modificó AD." -ForegroundColor Cyan
    Write-Log "FIN — MODO SIMULACIÓN. Total: $Script:Total | Sim: $Script:OK | Errores: $Script:Fail" -L SIM
} else {
    Write-Host "  ⚠️  Se realizaron cambios en Active Directory." -ForegroundColor Yellow
    Write-Host "  🔐 El Excel de resultados contiene contraseñas." -ForegroundColor Yellow
    Write-Host "     Elimínalo tras distribuir las credenciales." -ForegroundColor Yellow
    Write-Log "FIN — MODO REAL. Total: $Script:Total | OK: $Script:OK | Errores: $Script:Fail" -L OK
}
Write-Host ""
