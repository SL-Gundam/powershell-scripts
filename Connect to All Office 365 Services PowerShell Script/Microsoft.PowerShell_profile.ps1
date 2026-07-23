$Shell = $Host.UI.RawUI
$Shell.WindowTitle="SysadminGeek"

$GraphScopes = @(
    "User.Read.All"
)
$EntraScopes = @(
    "User.Read.All"
)
$CertificateValidityPeriod = (Get-Date).AddYears(1)

Function UpdateModules {
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

    if((($PSVersionTable::PSVersion.Major) -ge 7) -and (Get-InstalledModule -Name SharePointPnPPowerShellOnline -ErrorAction:SilentlyContinue))
    {
        # Uninstalling legacy module
        Write-Host Removing SharePointPnPPowerShellOnline. Legacy module Powershell 7.
        Uninstall-Module SharePointPnPPowerShellOnline -Force -AllVersions -ErrorAction:SilentlyContinue
    }

    if(Get-InstalledModule -Name AzureAD -ErrorAction:SilentlyContinue)
    {
        # Uninstalling legacy module
        Write-Host Removing AzureAD.
        Uninstall-Module AzureAD -Force -AllVersions -ErrorAction:SilentlyContinue
    }

    if(Get-InstalledModule -Name MSOnline -ErrorAction:SilentlyContinue)
    {
        # Uninstalling legacy module
        Write-Host Removing MSOnline.
        Uninstall-Module MSOnline -Force -AllVersions -ErrorAction:SilentlyContinue
    }

    $modules = Get-InstalledModule
    foreach ($module in $modules) {
        write-host Checking update for $module.Name 

        $lockModules = @(
            [pscustomobject]@{Name='ExchangeOnlineManagement';keepVersion='3.9.0'} # bug certificate auth for 3.10.0 for secandcompcenter Powershell 5 en 7 AND bug with MFA auth for 3.10.0 for secandcompcenter Powershell 7
            [pscustomobject]@{Name='Microsoft.Entra'; keepVersion='1.2.0'} # 1.3.0 still has login issues with credentials for other tenants
            [pscustomobject]@{Name='Microsoft.Graph'; keepVersion='2.33.0'} # 2.38.0 still has login issues with credentials for other tenants
            [pscustomobject]@{Name='Microsoft.Graph.Beta'; keepVersion='2.33.0'} # 2.38.0 still has login issues with credentials for other tenants
        )
        $versionLocked = $False

        foreach ( $lockModule in $lockModules ) {
            if ( $module.Name.StartsWith( $lockModule.Name ) ) {
                $versionLocked = $True

                if ( $module.Name -eq $lockModule.Name ) {
                    write-host Installing older module $lockModule.Name RequiredVersion $lockModule.keepVersion and uninstalling all other versions

                    # 1 Installeer gewenste versie
                    Install-Module $lockModule.Name -RequiredVersion $lockModule.keepVersion -AllowClobber -Scope CurrentUser

                    # Show latest version
                    Find-Module $lockModule.Name | Format-Table
                }

                $allVersions = Get-InstalledModule -Name $module.Name -AllVersions
    
                foreach ( $version in $allVersions ) {
                    write-host "Checking locked $($module.Name) $($version.Version)"
                    if ( $version.Version -ne $lockModule.keepVersion ) {
                        Write-Host "Removing $($module.Name) $($version.Version)"
                        Uninstall-Module -Name $module.Name -RequiredVersion $version.Version -Force
                    }
                }
            }
        }

        if ( $versionLocked -eq $False ) {
            Update-Module -Name $module.Name
        }
    }
    
    # Cleanup older versions of modules
    Write-Host "Cleanup older versions of modules"
    $modules = Get-InstalledModule
    foreach ($module in $modules) {
        Get-InstalledModule -Name $module.Name -AllVersions | Group-Object Name | ForEach-Object {
            $moduleGroup = $_.Group | Sort-Object { [version]$_.Version } -Descending
            $latest = $moduleGroup | Select-Object -First 1
            $olderVersions = $moduleGroup | Where-Object { $_.Version -ne $latest.Version }
            foreach ($old in $olderVersions) {
                Write-Host "Removing $($old.Name) v$($old.Version)"
                Uninstall-Module -Name $old.Name -RequiredVersion $old.Version -Force
            }
        }
    }

    # Check orphaned modules
    $profilemodulepath = (Split-Path $PROFILE) + "\Modules"
    $installed = Get-InstalledModule
    $modulePaths = @(
        "$env:ProgramFiles\WindowsPowerShell\Modules",
        $profilemodulepath
    )

    foreach ($path in $modulePaths) {
        Get-ChildItem -Path $path -Directory | ForEach-Object {
            if ($installed.Name -notcontains $_.Name) {
                Write-Host "Potential orphaned module: $($_.FullName)"
            }
        }
    }

    # Problems Sharepoint Online not updating
    if(Get-InstalledModule -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction:SilentlyContinue)
    {
        if((Get-InstalledModule -Name Microsoft.Online.SharePoint.PowerShell).Version -ne (Find-Module Microsoft.Online.SharePoint.PowerShell).Version)
        {
            # Uninstalling promatic module module
            Write-Host Removing Microsoft.Online.SharePoint.PowerShell to fix update issue.
            Uninstall-Module Microsoft.Online.SharePoint.PowerShell -Force -AllVersions -ErrorAction:SilentlyContinue
            $profilemodulepath = (Split-Path $PROFILE) + "\Modules"
            remove-item "$profilemodulepath\Microsoft.Online.SharePoint.PowerShell\*" -recurse
        }
    }
}

Function ConnectEXOnlineJSON {
    $JSON = Get-Content "$PSScriptRoot/Tenants.json" -ErrorAction SilentlyContinue | ConvertFrom-Json
    if (-not $JSON) {
        $JSON = @{}
    }

    if (-not $JSON.Tenants) {
        $JSON | Add-Member -MemberType NoteProperty -Name Tenants -Value (@{})
        $JSON.Tenants = @()
    }
    $tenants = $JSON.Tenants

    if($tenants.Count -eq 0)
    {
        Write-Host "No Tenants defined. Run JSONentries" -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }
    else
    {
        for ($i = 0; $i -lt $tenants.Count; $i++) {
            Write-Host "$($i+1). $($tenants[$i].Name)"
        }

        $choice = Read-Host "Select tenant"
        if ($choice -notmatch '^\d+$' -or [int]$choice -lt 1 -or [int]$choice -gt $tenants.Count) {
            Write-Host "Invalid selection" -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }
        $selectedTenant = $tenants[$choice - 1]


        Write-Host Available services: 'MSGraph','MSGraphBeta','MSTeams','SharePointOnline','SharePointPnP','SecAndCompCenter','ExchangeOnline','MSEntra'
        $Service=Read-Host -Prompt "Which service (leave empty for all)?"

        if($Service -eq ""){
            & $PSScriptRoot/ConnectO365Services.ps1 -TenantId $selectedTenant.TenantId -AppId $selectedTenant.AppId -CertificateThumbprint $selectedTenant.CertThumbprint
        }
        else{
            & $PSScriptRoot/ConnectO365Services.ps1 -TenantId $selectedTenant.TenantId -AppId $selectedTenant.AppId -CertificateThumbprint $selectedTenant.CertThumbprint -Services $Service
        }

        & GetCurrentAccounts
    }
}

Function ConnectEXOnlineMFA {
    Write-Host Available services: 'MSGraph','MSGraphBeta','MSTeams','SharePointOnline','SharePointPnP','SecAndCompCenter','ExchangeOnline','MSEntra'
    $Service=Read-Host -Prompt "Which service (leave empty for all)?"

    if($Service -eq ""){
        & $PSScriptRoot/ConnectO365Services.ps1 -MFA -GraphScopes $GraphScopes -EntraScopes $EntraScopes
    }
    else{
        & $PSScriptRoot/ConnectO365Services.ps1 -MFA -Services $Service -GraphScopes $GraphScopes -EntraScopes $EntraScopes
    }

    & GetCurrentAccounts
}

Function DisconnectEXOnline {
    & $PSScriptRoot/ConnectO365Services.ps1 -Disconnect

    Write-Host All disconnected
}

Function GetCurrentAccounts {
    Write-Host MSGraph: (Get-MgOrganization -ErrorAction:SilentlyContinue).DisplayName

    try {
        $GetEntraUser = (Get-EntraUser -Top 1 -ErrorAction:Stop).UserPrincipalName
        $Connected = $true
    }
    catch {
        $Connected = $false
    }
    write-host Entra: $GetEntraUser

    Write-Host Teams: (Get-CsTenant -ErrorAction:SilentlyContinue).DisplayName

    try {
        $GetSPOSite = (Get-SPOSite -Limit 1 -ErrorAction:Stop).Url
        $Connected = $true
    }
    catch {
        $Connected = $false
    }
    Write-Host SharePointOnline: $GetSPOSite

    try {
        $GetPnPConnection = (Get-PnPConnection -ErrorAction:Stop).Url
        $Connected = $true
    }
    catch {
        $Connected = $false
    }
    Write-Host SharePointPnP: $GetPnPConnection
    Write-Host "SecAndCompCenter & ExchangeOnline (returns 2 values if both connected): " (Get-ConnectionInformation -ErrorAction:SilentlyContinue).UserPrincipalName
}

function JSONentries{
    $JSON = Get-Content "$PSScriptRoot/Tenants.json" -ErrorAction SilentlyContinue | ConvertFrom-Json
    if (-not $JSON) {
        $JSON = @{}
    }

    if (-not $JSON.Tenants) {
        $JSON | Add-Member -MemberType NoteProperty -Name Tenants -Value (@{})
        $JSON.Tenants = @()
    }

    # Haal huidige tenant op
    $Org = Get-MgOrganization -ErrorAction SilentlyContinue
    $CurrentTenantName = $Org.DisplayName

    # Zoek bestaande entry op tenantnaam
    $ExistingTenantIndex = -1

    for ($i = 0; $i -lt $JSON.Tenants.Count; $i++) {
        if ($JSON.Tenants[$i].Name -eq $CurrentTenantName) {
            $ExistingTenantIndex = $i
            break
        }
    }

    Write-Host ""
    Write-Host "Current tenant: $CurrentTenantName" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. New entry"


    if ($ExistingTenantIndex -ge 0) {
        Write-Host "2. Overwrite existing entry"
    }
    else {
        Write-Host "2. Overwrite existing entry - unavailable, no matching tenant name found" -ForegroundColor DarkGray
    }

    Write-Host "3. Delete entry"
    Write-Host ""

    $choice = Read-Host "Select action"

    if ($choice -notmatch '^[1-3]$') {
        Write-Host "Invalid selection" -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }

    if ($choice -eq '2' -and $ExistingTenantIndex -lt 0) {
        Write-Host "Overwrite is not available because no existing JSON entry matches tenant name '$CurrentTenantName'." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return
    }

    if ($choice -eq '3') {
        if (-not $JSON.Tenants -or $JSON.Tenants.Count -eq 0) {
            Write-Host "No JSON entries found to delete." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            return
        }

        Write-Host ""
        Write-Host "Select entry to delete:" -ForegroundColor Cyan

        for ($i = 0; $i -lt $JSON.Tenants.Count; $i++) {
            Write-Host "$($i + 1). $($JSON.Tenants[$i].Name) - $($JSON.Tenants[$i].TenantId)"
        }

        Write-Host "0. Cancel"
        Write-Host ""

        $deleteChoice = Read-Host "Select entry"

        if ($deleteChoice -eq '0') {
            Write-Host "Delete cancelled" -ForegroundColor Yellow
            return
        }

        if (
            $deleteChoice -notmatch '^\d+$' -or
            [int]$deleteChoice -lt 1 -or
            [int]$deleteChoice -gt $JSON.Tenants.Count
        ) {
            Write-Host "Invalid selection" -ForegroundColor Red
            Start-Sleep -Seconds 2
            return
        }

        $deleteIndex = [int]$deleteChoice - 1
        $deletedEntry = $JSON.Tenants[$deleteIndex]

        $JSON.Tenants = @(
            for ($i = 0; $i -lt $JSON.Tenants.Count; $i++) {
                if ($i -ne $deleteIndex) {
                    $JSON.Tenants[$i]
                }
            }
        )

        $JSON | ConvertTo-Json | Set-Content -Path "$PSScriptRoot/Tenants.json"

        Write-Host "Deleted JSON entry: $($deletedEntry.Name)" -ForegroundColor Yellow
        return
    }


    # Start create application

    $cert = Get-ChildItem -Path Cert:\CurrentUser\My\ | where { $_.subject -eq "CN=$($Shell.WindowTitle)" } -ErrorAction SilentlyContinue
    if (-not $cert)
    {
        $cert = New-SelfSignedCertificate `
            -Subject "CN=$($Shell.WindowTitle)" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -KeySpec KeyExchange `
            -KeyLength 2048 `
            -NotAfter $CertificateValidityPeriod

        Write-Host "Certificate generated" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Existing certificate found" -ForegroundColor Yellow
    }
    Export-Certificate -Cert $cert -FilePath $PSScriptRoot\$($Shell.WindowTitle).cer

    $CertThumb = $cert.Thumbprint


    # Create the application with the necessary permissions
    $app = Get-MgApplication -ConsistencyLevel eventual -Filter "DisplayName eq '$($Shell.WindowTitle)'"
    if (-not $app) {
        $app = New-MgApplication -DisplayName "$($Shell.WindowTitle)"
        $sp = New-MgServicePrincipal -AppId $app.AppId

        Write-Host "New App created" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Existing App found" -ForegroundColor Yellow
    }

    # Check whether certificate already exists
    $existingCert = $app.KeyCredentials | Where-Object {
        $_.CustomKeyIdentifier -and
        ([System.Convert]::ToBase64String($_.CustomKeyIdentifier).ToUpperInvariant() -eq $CertThumb.ToUpperInvariant())
    }

    if ($existingCert) {
        Write-Host "Certificate already exists on app registration" -ForegroundColor Yellow
    }
    else {
        $newKeyCredential = @{
            Type          = "AsymmetricX509Cert"
            Usage         = "Verify"
            Key           = $cert.RawData
            DisplayName   = "$Env:ComputerName $([Environment]::UserName)"
            StartDateTime = $cert.NotBefore
            EndDateTime   = $cert.NotAfter
        }

        $updatedKeyCredentials = @($app.KeyCredentials) + $newKeyCredential

        Update-MgApplication `
            -ApplicationId $app.Id `
            -KeyCredentials $updatedKeyCredentials
        Write-Host "Certificate imported to App" -ForegroundColor Yellow
    }


    $requiredResourceAccess = @()
    $graphResourceAccess = @()
    $exoResourceAccess = @()
    $spoResourceAccess = @()

    # Get existing service principals
    $appSp = Get-MgServicePrincipal `
        -Filter "appId eq '$($app.AppId)'"

    $appSpAra = Get-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $appSp.Id

    # Graph service principal
    $graphSp = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

    $permissions = $GraphScopes

    foreach ($permission in $permissions) {
        $appRole = $graphSp.AppRoles |
            Where-Object {
                $_.Value -eq $permission -and
                $_.AllowedMemberTypes -contains "Application"
            }

        $existingAssignment = $appSpAra |
            Where-Object {
                $_.ResourceId -eq $graphSp.Id -and
                $_.AppRoleId -eq $appRole.Id
            }

        $graphResourceAccess += @{
            Id   = $appRole.Id
            Type = "Role"
        }
 
        if (-not $existingAssignment) {
            $dummy = New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $appSp.Id `
                -PrincipalId $appSp.Id `
                -ResourceId $graphSp.Id `
                -AppRoleId $appRole.Id

            Write-Host "New Graph permission added: $permission" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "Graph permission already exists: $permission" -ForegroundColor Yellow
        }
    }

    $requiredResourceAccess += @{
        ResourceAppId  = $graphSp.AppId
        ResourceAccess = @($graphResourceAccess)
    }


    # Exchange service principal
    $exoSp = Get-MgServicePrincipal `
        -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'"

    $permissions = @(
#        "full_access_as_app", # old EXO permission. Should not be needed
        "Exchange.AdminAPI.ManageAsApp",
        "Exchange.ManageAsApp",
        "Exchange.ManageAsAppV2"
    )

    foreach ($permission in $permissions) {
        $appRole = $exoSp.AppRoles |
            Where-Object {
                $_.Value -eq $permission -and
                $_.AllowedMemberTypes -contains "Application"
            }

        $existingAssignment = $appSpAra |
            Where-Object {
                $_.ResourceId -eq $exoSp.Id -and
                $_.AppRoleId -eq $appRole.Id
            }

        $exoResourceAccess += @{
            Id   = $appRole.Id
            Type = "Role"
        }
 
        if (-not $existingAssignment) {
            $dummy = New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $appSp.Id `
                -PrincipalId $appSp.Id `
                -ResourceId $exoSp.Id `
                -AppRoleId $appRole.Id

            Write-Host "New EXO permission added: $permission" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "EXO permission already exists: $permission" -ForegroundColor Yellow
        }
    }

    $requiredResourceAccess += @{
        ResourceAppId  = $exoSp.AppId
        ResourceAccess = @($exoResourceAccess)
    }


    # Sharepoint service principal
    $spoSp = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'"

    $permissions = @(
        "Sites.FullControl.All",
#        "AllSites.FullControl", # Delegated permission not available for applications
        "TermStore.ReadWrite.All",
        "User.ReadWrite.All"
    )

    foreach ($permission in $permissions) {
        $appRole = $spoSp.AppRoles |
            Where-Object {
                $_.Value -eq $permission -and
                $_.AllowedMemberTypes -contains "Application"
            }

        $existingAssignment = $appSpAra |
            Where-Object {
                $_.ResourceId -eq $spoSp.Id -and
                $_.AppRoleId -eq $appRole.Id
            }

        $spoResourceAccess += @{
            Id   = $appRole.Id
            Type = "Role"
        }

        if (-not $existingAssignment) {
            $dummy = New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $appSp.Id `
                -PrincipalId $appSp.Id `
                -ResourceId $spoSp.Id `
                -AppRoleId $appRole.Id

            Write-Host "New Sharepoint permission added: $permission" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "Sharepoint permission already exists: $permission" -ForegroundColor Yellow
        }
    }

    $requiredResourceAccess += @{
        ResourceAppId  = $spoSp.AppId
        ResourceAccess = @($spoResourceAccess)
    }


    # Write configured permissions to app registration
    Update-MgApplication `
        -ApplicationId $app.Id `
        -RequiredResourceAccess $requiredResourceAccess
    Write-Host "App updated with configured permissions" -ForegroundColor Yellow


    # Get Global Administrator role
    $role = Get-MgDirectoryRole |
        Where-Object { $_.DisplayName -eq "Global Administrator" }

    # Check if service principal is already a member
    $existingMember = Get-MgDirectoryRoleMemberAsServicePrincipal -DirectoryRoleId $role.Id -All |
        Where-Object { $_.Id -eq $appSp.Id }

    if (-not $existingMember) {
        Write-Host "Adding service principal to Global Administrator role..." -ForegroundColor Yellow

        New-MgDirectoryRoleMemberByRef `
            -DirectoryRoleId $role.Id `
            -BodyParameter @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($appSp.Id)"
            }
    }
    else {
        Write-Host "Service principal is already a member of the Global Administrator role." -ForegroundColor Yellow
    }

    # End create application

    $Entry = @{
        Name = $Org.DisplayName
        TenantId = $Org.Id
        AppId = $app.AppId
        CertThumbprint = $CertThumb
    }
    if ($choice -eq '1' -or $ExistingTenantIndex -lt 0)
    {
        $JSON.Tenants += $Entry
        Write-Host "JSON entry added" -ForegroundColor Yellow
    }
    else{
        $JSON.Tenants[$ExistingTenantIndex] = $Entry
        Write-Host "JSON entry modified: $($Org.DisplayName)" -ForegroundColor Yellow
    }

    $JSON | ConvertTo-Json | Set-Content -Path "$PSScriptRoot/Tenants.json"
}

function Reload-Profile {
    Write-Host "Reloading PowerShell profile from $PROFILE" -ForegroundColor Cyan
    . $PROFILE
}

function Show-FunctionMenu {
    $filePath = $PSCommandPath
    $functionNames = @()

    if (Test-Path $filePath) {
        $content = Get-Content $filePath -Raw
        $matches = [regex]::Matches($content, '(?im)^\s*function\s+([a-zA-Z0-9_-]+)')

        foreach ($match in $matches) {
            $functionNames += $match.Groups[1].Value
        }
    }

    if (-not $functionNames) {
        Write-Host "No functions found in profile script." -ForegroundColor Red
        return
    }
    else
    {
        $functionNames = $functionNames | Where-Object { $_ -ne "Show-FunctionMenu" }
        $functionNames = $functionNames | Where-Object { $_ -ne "Test-NoPowerShellArguments" }

        $functionNames += "Exit"
    }

    while ($true) {
        Clear-Host
        Write-Host "==== PowerShell Function Menu ====" -ForegroundColor Cyan
        Write-Host "Select a function to run:`n"

        for ($i = 0; $i -lt $functionNames.Count; $i++) {
            Write-Host "$($i + 1). $($functionNames[$i])"
        }

        $selection = Read-Host "`nEnter your choice (number)"
        if ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $functionNames.Count) {
            $choice = $functionNames[$selection - 1]

            if ($choice -eq 'Exit') {
                break
            }

            #try {
                Write-Host "`nRunning function: $choice" -ForegroundColor Yellow
                & $choice
            #}
            #catch {
            #    Write-Error "Error executing function: $_"
            #}

            Write-Host "`nPress Enter to return to the menu..."
            [void][System.Console]::ReadLine()
        }
        else {
            Write-Host "Invalid selection. Try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }

        $functionNames = $functionNames
    }
}

function Test-NoPowerShellArguments {
    $cmd = [Environment]::CommandLine.Trim()

    if ([string]::IsNullOrWhiteSpace($cmd)) {
        return $false
    }

    # Remove first token: quoted executable path or unquoted executable path
    if ($cmd -match '^\s*"[^"]+"\s*(?<args>.*)$') {
        $remaining = $Matches.args.Trim()
    }
    elseif ($cmd -match '^\s*\S+\s*(?<args>.*)$') {
        $remaining = $Matches.args.Trim()
    }
    else {
        return $false
    }

    if ($remaining -eq "-WorkingDirectory ~") {
        return $true
    }

    return [string]::IsNullOrWhiteSpace($remaining)
}

if (Test-NoPowerShellArguments) {
    Show-FunctionMenu
}
