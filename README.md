# PowerShell AD Tools

Conjunto de scripts PowerShell para administración masiva de usuarios en Active Directory, orientados a entornos Windows Server con cargas de trabajo empresariales o gubernamentales.

---

## Scripts incluidos

### `Crear-UsuariosAD.ps1`
Crea usuarios en Active Directory de forma masiva a partir de un archivo Excel.

**Características:**
- Lee columnas: Nombre, Apellido, Usuario, Departamento, Compania, Calle, Ciudad, Estado, CP, Pais, Email, Telefono
- Genera contraseñas seguras con `RandomNumberGenerator` (.NET Cryptography)
- Crea automáticamente la OU si no existe
- Valida duplicados de SAMAccountName y UPN antes de crear
- Modo simulación (`-Simular`) para dry-run sin cambios en AD
- Exporta resultados en Excel (hoja Usuarios Creados + hoja Errores)
- Log detallado con timestamp por ejecución

**Uso:**
```powershell
# Modo real
.\Crear-UsuariosAD.ps1 -ExcelEntrada ".\usuarios.xlsx" -ExcelSalida ".\resultados.xlsx" -DominioUPN "empresa.dominio.com"

# Modo simulación
.\Crear-UsuariosAD.ps1 -ExcelEntrada ".\usuarios.xlsx" -ExcelSalida ".\resultados.xlsx" -DominioUPN "empresa.dominio.com" -Simular
```

---

### `Actualizar-UsuariosAD.ps1`
Actualiza masivamente cargos, OUs y número de empleado en Active Directory desde un archivo Excel proporcionado por RH.

**Características:**
- Lee columnas: # Empleado, Nombre Completo, Unidad Administrativa, Cargo
- Busca usuarios por DisplayName con y sin acentos (4 estrategias de búsqueda)
- Mueve usuarios entre OUs y actualiza Title y EmployeeNumber
- Detecta y omite registros sin cambios reales
- Crea la OU destino automáticamente si no existe
- Modo simulación (`-Simular`) para dry-run sin cambios en AD
- Exporta resultados en Excel (3 hojas: Actualizados, Sin Cambios, Errores)
- Log detallado con timestamp por ejecución

**Uso:**
```powershell
# Modo real
.\Actualizar-UsuariosAD.ps1 -ExcelEntrada ".\rh_cargos.xlsx" -ExcelSalida ".\resultados.xlsx"

# Modo simulación
.\Actualizar-UsuariosAD.ps1 -ExcelEntrada ".\rh_cargos.xlsx" -ExcelSalida ".\resultados.xlsx" -Simular
```

---

## Requisitos

- Windows Server con PowerShell 5.1 o superior
- Módulo `ActiveDirectory` (RSAT): `Install-WindowsFeature RSAT-AD-PowerShell`
- Módulo `ImportExcel`: instalado automáticamente si no está presente

---

## Notas de seguridad

- El Excel de resultados de creación contiene contraseñas en texto plano. Distribúyelo de forma segura y elimínalo una vez entregadas las credenciales.
- Ningún script contiene credenciales, IPs ni datos de dominio hardcodeados.
- Se recomienda ejecutar siempre con `-Simular` primero para validar antes de aplicar cambios reales.

---

## Autor

[omar-gopar](https://github.com/omar-gopar) — Ingeniero de Infraestructura y Sistemas
