# Creates N screenings in Dynamics, then adds a row per screening to the
# TriggerCopyDataverseItemToSharePoint SharePoint list.
#
# Requires: Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser
# (1.12.0 specifically - newer PnP.PowerShell versions need PowerShell 7+, this needs to
# run on Windows PowerShell 5.1).
#
# Example usage:
# .\Create-Screenings-and-Trigger-Flow.ps1 -EnvironmentUrl "https://p365fedscreenqa.crm9.dynamics.com" `
#     -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" `
#     -Count 1000 -BatchSize 100

param(
    [Parameter(Mandatory)]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [string]$SharePointSiteUrl = 'https://frbprod1.sharepoint.com/sites/SYS-FDR/STAGE',

    [string]$SharePointListName = 'TriggerCopyDataverseItemToSharePoint',

    # SumRptTemplateID value written to every SharePoint row (fixed for now, may change later).
    [int]$SumRptTemplateId = -1,

    [string]$EntityPath = 'cr9da_fsn_screeningses',

    [int]$Count = 1000,
    [int]$BatchSize = 100,
    [int]$StartIndex = 1,

    [string]$CandidateFirstNamePrefix = 'LoadTest',
    [string]$CandidateLastNamePrefix = 'Candidate',
    [string]$CandidateEmailDomain = 'example.com',

    [string]$ResultsDir = "$PSScriptRoot\results"
)

# Checks before running script

if (-not ([System.Uri]::IsWellFormedUriString($EnvironmentUrl, [System.UriKind]::Absolute))) {
    Write-Error "EnvironmentUrl '$EnvironmentUrl' is not a valid URL."
    exit 1
}

if (-not ([System.Uri]::IsWellFormedUriString($SharePointSiteUrl, [System.UriKind]::Absolute))) {
    Write-Error "SharePointSiteUrl '$SharePointSiteUrl' is not a valid URL."
    exit 1
}

if ($Count -le 0) {
    Write-Error '-Count must be greater than 0.'
    exit 1
}

if ($BatchSize -le 0) {
    Write-Error '-BatchSize must be greater than 0.'
    exit 1
}

if ($BatchSize -gt $Count) {
    $BatchSize = $Count
}

# Fields common to every screening. Per-record values (name/email) are overridden
# per iteration below so 1000+ runs don't all create identical records.
$staticScreeningFields = @{
    cr9da_acknowledgebackground        = 'Yes'
    cr9da_acknowledgecitizenship       = 'Yes'
    cr9da_acknowledgeoffer             = 'Yes'
    cr9da_acknowledgeotherdocs         = 'Yes'
    cr9da_activeinactive               = 'true'
    cr9da_businessline                 = 'Credit Risk Management'
    cr9da_citizenship                  = 'U.S. Citizen only'
    cr9da_creditcheck                  = 'true'
    cr9da_dmv                          = 'false'
    cr9da_drugscreening                = 'false'
    cr9da_expedite                     = 'false'
    cr9da_fbifingerprinting            = 'true'
    cr9da_hiringmanagerdisplayname     = 'Traband, Lester J'
    cr9da_hiringmanageremail           = 'Lester.Traband@phil.frb.org'
    cr9da_initiatordisplayname         = 'Fedalen, Connor'
    cr9da_initiatoremail               = 'Connor.Fedalen@phil.frb.org'
    cr9da_initiatorrole                = 'FedScreen Coordinator, FedScreen Screening Partner, '
    cr9da_marijuana                    = 'false'
    cr9da_mediascreening               = 'false'
    cr9da_middlename                   = ''
    cr9da_otherscreeningreason         = ''
    cr9da_pdqidentifier                = 'Treasury High'
    cr9da_provisionalstart             = 'false'
    cr9da_requestedbydisplayname       = 'Fedalen, Connor'
    cr9da_requestedbyemail             = 'Connor.Fedalen@phil.frb.org'
    cr9da_requestingdistrict           = 'Philadelphia'
    cr9da_safrscreening                = 'Tier II - Credit Check Enhanced'
    cr9da_screeningreason              = 'New Hire'
    cr9da_screeningstatus              = 'In Progress'
    cr9da_spattestation                = 'Yes, I have confirmed that the submitted screening request matches the position screening requirements in Workday.'
    cr9da_spattestationdate            = (Get-Date -Format 'yyyy-MM-dd')
    cr9da_spattestedby                 = 'Connor Fedalen'
    cr9da_socialmedia                  = 'false'
    cr9da_startdate                    = (Get-Date).AddDays(14).ToString('yyyy-MM-dd')
    cr9da_submissiondate               = (Get-Date -Format 'yyyy-MM-ddT00:00:00')
    cr9da_submitterattestation         = 'Yes, the requested screening matches the Workday requirements.'
    cr9da_submitterattestationcomments = ''
    cr9da_treasury                     = 'true'
    cr9da_vendorscreeningpackages      = 'New Hire/CWR + Credit'
    cr9da_workdayposition              = '999999'
    cr9da_workertype                   = 'Employee'
}

# Authentication

$scope = $EnvironmentUrl.TrimEnd('/') + '/.default'

$deviceResponse = Invoke-RestMethod `
    -Method POST `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = "$scope offline_access" } `
    -ContentType 'application/x-www-form-urlencoded'

Write-Host $deviceResponse.message
Write-Host ''

$tokenBody = @{
    grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
    client_id   = $ClientId
    device_code = $deviceResponse.device_code
}

$accessToken = $null
$refreshToken = $null
$interval = [int]$deviceResponse.interval
while ($null -eq $accessToken) {
    Start-Sleep -Seconds $interval
    try {
        $tokenResponse = Invoke-RestMethod `
            -Method POST `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body $tokenBody `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop
        $accessToken = $tokenResponse.access_token
        $refreshToken = $tokenResponse.refresh_token
    }
    catch {
        $err = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue).error
        if ($err -eq 'authorization_pending') { continue }
        elseif ($err -eq 'slow_down') { $interval += 5; continue }
        elseif ($err -eq 'authorization_declined') { Write-Error 'Login was declined.'; exit 1 }
        elseif ($err -eq 'expired_token') { Write-Error 'Device code expired.'; exit 1 }
        else { Write-Error $_; exit 1 }
    }
}

Write-Host 'Authenticated successfully.'

# SharePoint auth is separate: cookie-based web login as the signed-in user, not tied to
# this app registration's Graph/API permissions (avoids needing tenant admin consent).
# Each runspace creates its own SharePoint connection so items in the same batch can run concurrently
# without sharing one CSOM client context across threads.
Import-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -ErrorAction Stop
Write-Host 'SharePoint auth will be opened inside each worker for concurrent batch execution.'

# Create screenings in batches; each item creates its Dynamics record and adds its SharePoint row
# inside the same worker so all items in a batch run concurrently without deadlocking.
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$runDir = Join-Path $ResultsDir $timestamp
New-Item -ItemType Directory -Path $runDir -Force | Out-Null
$csvPath = Join-Path $runDir 'screening-sharepoint-results.csv'

$createUri = $EnvironmentUrl.TrimEnd('/') + "/api/data/v9.0/$EntityPath"
$indices = $StartIndex..($StartIndex + $Count - 1)
$batches = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $indices.Count; $i += $BatchSize) {
    $batches.Add(@($indices[$i..([math]::Min($i + $BatchSize, $indices.Count) - 1)]))
}

# Runspace pools (not ForEach-Object -Parallel) so this also runs on Windows PowerShell 5.1.
$workerScript = {
    param($index, $accessToken, $createUri, $sharePointSiteUrl, $sharePointListName, $sumRptTemplateId, $firstNamePrefix, $lastNamePrefix, $emailDomain, $staticScreeningFields)

    Import-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -ErrorAction Stop

    $firstName = "$firstNamePrefix$index"
    $lastName = "$lastNamePrefix$index"

    $body = $staticScreeningFields.Clone()
    $body.cr9da_firstname = $firstName
    $body.cr9da_lastname = $lastName
    $body.cr9da_name = "$firstName $lastName"
    $body.cr9da_candidateemail = "$($firstName.ToLower())@$emailDomain"

    $result = [pscustomobject]@{
        Index             = $index
        FirstName         = $firstName
        LastName          = $lastName
        ScreeningGuid     = $null
        CreateSuccess     = $false
        CreateError       = $null
        SharePointSuccess = $false
        SharePointError   = $null
    }

    try {
        $createResponse = Invoke-RestMethod `
            -Method POST `
            -Uri $createUri `
            -Headers @{
            Authorization      = "Bearer $accessToken"
            Accept             = 'application/json'
            'OData-MaxVersion' = '4.0'
            'OData-Version'    = '4.0'
            Prefer             = 'return=representation'
        } `
            -ContentType 'application/json' `
            -Body ($body | ConvertTo-Json -Depth 5) `
            -ErrorAction Stop

        $result.ScreeningGuid = $createResponse.cr9da_fsn_screeningsid
        $result.CreateSuccess = $true
    }
    catch {
        $result.CreateError = $_.Exception.Message
        return $result
    }

    try {
        $sharePointConnection = Connect-PnPOnline -Url $sharePointSiteUrl -UseWebLogin -ReturnConnection -WarningAction SilentlyContinue
        Add-PnPListItem `
            -List $sharePointListName `
            -Values @{
            Title            = "$lastName, $firstName"
            FSN_ScreeningID  = $result.ScreeningGuid
            SumRptTemplateID = $sumRptTemplateId
        } `
            -Connection $sharePointConnection `
            -ErrorAction Stop | Out-Null
        $result.SharePointSuccess = $true
        Disconnect-PnPOnline -Connection $sharePointConnection -ErrorAction SilentlyContinue
    }
    catch {
        $result.SharePointError = $_.Exception.Message
    }

    return $result
}

$allResults = [System.Collections.Generic.List[object]]::new()
$batchNum = 0

foreach ($batch in $batches) {
    $batchNum++
    Write-Host "Batch $batchNum of $($batches.Count) ($($batch.Count) screenings)..."

    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $BatchSize)
    $runspacePool.Open()

    $jobs = foreach ($index in $batch) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool
        [void]$ps.AddScript($workerScript).AddArgument($index).AddArgument($accessToken).AddArgument($createUri).AddArgument($SharePointSiteUrl).AddArgument($SharePointListName).AddArgument($SumRptTemplateId).AddArgument($CandidateFirstNamePrefix).AddArgument($CandidateLastNamePrefix).AddArgument($CandidateEmailDomain).AddArgument($staticScreeningFields)
        [pscustomobject]@{
            PowerShell = $ps
            Handle     = $ps.BeginInvoke()
        }
    }

    $batchResults = foreach ($job in $jobs) {
        $job.PowerShell.EndInvoke($job.Handle)
        $job.PowerShell.Dispose()
    }

    $runspacePool.Close()
    $runspacePool.Dispose()

    $allResults.AddRange(@($batchResults))

    $createdInBatch = @($batchResults | Where-Object CreateSuccess).Count
    $addedInBatch = @($batchResults | Where-Object SharePointSuccess).Count
    Write-Host "  Created: $createdInBatch/$($batch.Count), Added to SharePoint: $addedInBatch/$($batch.Count)"
}

$allResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$totalCreated = @($allResults | Where-Object CreateSuccess).Count
$totalAdded = @($allResults | Where-Object SharePointSuccess).Count
$totalFailed = $allResults.Count - $totalAdded

Write-Host ''
Write-Host "Done. Screenings created: $totalCreated/$($allResults.Count). Added to SharePoint: $totalAdded/$($allResults.Count)."
if ($totalFailed -gt 0) {
    Write-Warning "$totalFailed record(s) had a create or SharePoint failure. See $csvPath for per-record details."
}
else {
    Write-Host "Results saved to: $csvPath"
}
