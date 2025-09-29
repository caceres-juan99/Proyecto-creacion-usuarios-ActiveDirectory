Import-Module ActiveDirectory

# Ruta fija del archivo CSV con los usuarios
$Ruta = "C:\Users\jcaceres\OneDrive - Cusezar S.A\Cusezar S.A\SCRIPTs\PYTHON - Crear usuarios nuevos en AD\Create CSV New users - Power Automate\Input_script.csv"

# Leer y procesar el CSV
Import-Csv -Path $Ruta -Delimiter ";" | ForEach-Object {
    if (![string]::IsNullOrWhiteSpace($_.Contraseña)) {
        try {
            New-ADUser `
                -Name $_.NombreCompleto `
                -SamAccountName $_.Usuario `
                -UserPrincipalName ($_.Usuario + "@cusezar.com") `
                -GivenName $_.Nombres `
                -Surname $_.Apellidos `
                -Initials $_.Iniciales `
                -DisplayName $_.NombreCompleto `
                -EmailAddress $_.Correo `
                -Office $_.Oficina `
                -Title $_.Puesto `
                -Department $_.Departamento `
                -AccountPassword (ConvertTo-SecureString $_.Contraseña -AsPlainText -Force) `
                -Path $_.OU `
                -Company $_.Organizacion `
                -Enabled $true `
                -ChangePasswordAtLogon $false `
                -OtherAttributes @{ proxyAddresses = ($_.proxyAddresses -split ";") } `
                -Verbose
        } catch {
            Write-Warning "⚠️ Error creando el usuario $($_.Usuario): $_"
        }
    } else {
        Write-Warning "⚠️ Usuario $($_.Usuario) no tiene contraseña. No se creó."
    }
}

Write-Host "`n✅ La creación de usuarios ha finalizado correctamente."