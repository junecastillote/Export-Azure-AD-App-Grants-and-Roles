[CmdletBinding(DefaultParameterSetName = 'Default')]
param (
    [Parameter(ParameterSetName = 'ByServicePrincipalId')]
    [ValidateNotNullOrEmpty()]
    [string]
    $ServicePrincipalId,

    [Parameter(ParameterSetName = 'ByDisplayName')]
    [ValidateNotNullOrEmpty()]
    [string]
    $DisplayName,

    [Parameter(ParameterSetName = 'Default')]
    [switch]
    $All
)

switch ($PSCmdlet.ParameterSetName) {
    ByServicePrincipalId {
        # Get specific service pricipal object by ID
        try {
            $all_sp = @(Get-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId -ErrorAction Stop)
        }
        catch {
            Write-Error $_.Exception.Message
            return $null
        }
    }
    ByDisplayName {
        # Get specific service pricipal object by displayname
        try {
            $all_sp = @(Get-MgServicePrincipal -Filter "DisplayName eq '$($DisplayName)'" -ErrorAction Stop)
        }
        catch {
            Write-Error $_.Exception.Message
            return $null
        }
    }
    Default {
        # Get all service principal objects
        Write-Verbose "Getting all service principal objects..."
        $all_sp = @(Get-MgServicePrincipal -All)
    }
}

if ($all_sp.Count -lt 1) {
    return $null
}

$total = $all_sp.Count

for ($i = 0 ; $i -lt $total ; $i++) {
    # Calculate the percentage completed
    $percentComplete = [math]::Round(($i / $total) * 100, 2)

    # Display the progress bar
    Write-Progress -Activity "Processing Service Principals [$($i+1) / $total]" `
        -Status "Processing: $($all_sp[$i].DisplayName)" `
        -PercentComplete $percentComplete

    # Get the delegated permissions
    $delegated_permisions = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $all_sp[$i].Id

    # Below are the scopes that permits delegated access to OneDrive and SharePoint sites/items (least to most privilege per REST API)
    $reference_scopes = @(
        'MyFiles.Read', # Start: SharePoint REST API
        'MyFiles.Write',
        'AllSites.Read',
        'AllSites.Manage',
        'AllSites.Write',
        'AllSites.FullControl',
        'Files.Read', # Start: Microsoft Graph API
        'Files.ReadWrite',
        'Files.Read.All',
        'Files.ReadWrite.All',
        'Group.Read.All',
        'Group.ReadWrite.All',
        'Sites.Read.All',
        'Sites.ReadWrite.All'
    )

    # Filter permissions for OneDrive
    $current_scopes = @(($delegated_permisions | Select-Object -Unique Scope).Scope)
    if ($current_scopes.Count -lt 0) {
        Continue
    }
    $matching_scopes = @($current_scopes | Where-Object { $_ -in $reference_scopes })

    if ($matching_scopes.Count -lt 1) {
        Continue
    }

    $spo_odb_permissions = ($matching_scopes | Select-Object -Unique) -join "," -replace " ", ","
    [PSCustomObject]@{
        DisplayName          = $all_sp[$i].DisplayName
        Id                   = $all_sp[$i].Id
        SignInAudience       = $all_sp[$i].SignInAudience
        ServicePrincipalType = $all_sp[$i].ServicePrincipalType
        Scopes               = $spo_odb_permissions
    }
}