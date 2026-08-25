# Creates N screenings in Dynamics, then uses each screening's GUID to trigger a
# Power Automate flow via the run-only apihub trigger endpoint (JSON body: {text, number}).
#
# Example usage:
# .\Create-Screenings-and-Trigger-Flow.ps1 -EnvironmentUrl "https://p365fedscreenqa.crm9.dynamics.com" `
#     -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" `
#     -FlowUrl "https://power-apis-usgov001-public.azure-apihub.us/apim/logicflows/<flow-guid>/triggers/manual/run?api-version=2016-11-01" `
#     -Count 1000 -BatchSize 100

param(
    [Parameter(Mandatory)]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    # Power Automate run-only apihub trigger URL (from DevTools, not the flow details page link).
    [Parameter(Mandatory)]
    [string]$FlowUrl,

    # Second field the trigger's JSON body expects alongside the screening GUID ({"text":<guid>,"number":FlowNumberValue}).
    [int]$FlowNumberValue = -1,

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

if (-not ([System.Uri]::IsWellFormedUriString($FlowUrl, [System.UriKind]::Absolute))) {
    Write-Error "FlowUrl '$FlowUrl' is not a valid URL."
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

# The apihub trigger endpoint is a different resource than Dynamics, so it needs its own token
# (audience = the apihub host itself, discovered from a captured browser request). Reuse the
# refresh token from the first login instead of prompting for a second interactive sign-in.
$flowScope = ([System.Uri]$FlowUrl).GetLeftPart([System.UriPartial]::Authority) + '/.default'

if (-not $refreshToken) {
    Write-Error 'No refresh token was returned from the first sign-in, so a token for the flow endpoint cannot be obtained silently.'
    exit 1
}

$flowTokenBody = @{
    grant_type    = 'refresh_token'
    client_id     = $ClientId
    refresh_token = $refreshToken
    scope         = $flowScope
}

$flowTokenResponse = Invoke-RestMethod `
    -Method POST `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Body $flowTokenBody `
    -ContentType 'application/x-www-form-urlencoded' `
    -ErrorAction Stop

$flowAccessToken = $flowTokenResponse.access_token

Write-Host 'Obtained flow trigger token silently via refresh token.'

# Create screenings and trigger the flow, in batches

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$runDir = Join-Path $ResultsDir $timestamp
New-Item -ItemType Directory -Path $runDir -Force | Out-Null
$csvPath = Join-Path $runDir 'screening-flow-results.csv'

$createUri = $EnvironmentUrl.TrimEnd('/') + "/api/data/v9.0/$EntityPath"
$indices = $StartIndex..($StartIndex + $Count - 1)
$batches = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $indices.Count; $i += $BatchSize) {
    $batches.Add(@($indices[$i..([math]::Min($i + $BatchSize, $indices.Count) - 1)]))
}

# Runspace pools (not ForEach-Object -Parallel) so this also runs on Windows PowerShell 5.1.
$workerScript = {
    param($index, $accessToken, $createUri, $flowUrl, $flowAccessToken, $flowNumberValue, $firstNamePrefix, $lastNamePrefix, $emailDomain, $staticScreeningFields)

    $firstName = "$firstNamePrefix$index"
    $lastName = "$lastNamePrefix$index"

    $body = $staticScreeningFields.Clone()
    $body.cr9da_firstname = $firstName
    $body.cr9da_lastname = $lastName
    $body.cr9da_name = "$firstName $lastName"
    $body.cr9da_candidateemail = "$($firstName.ToLower())@$emailDomain"

    $result = [pscustomobject]@{
        Index         = $index
        ScreeningGuid = $null
        CreateSuccess = $false
        CreateError   = $null
        FlowSuccess   = $false
        FlowError     = $null
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
        $flowBody = @{ text = $result.ScreeningGuid; number = $flowNumberValue } | ConvertTo-Json
        Invoke-RestMethod -Method POST -Uri $flowUrl -Headers @{ Authorization = "Bearer $flowAccessToken" } -ContentType 'application/json' -Body $flowBody -ErrorAction Stop | Out-Null
        $result.FlowSuccess = $true
    }
    catch {
        $result.FlowError = $_.Exception.Message
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
        [void]$ps.AddScript($workerScript).AddArgument($index).AddArgument($accessToken).AddArgument($createUri).AddArgument($FlowUrl).AddArgument($flowAccessToken).AddArgument($FlowNumberValue).AddArgument($CandidateFirstNamePrefix).AddArgument($CandidateLastNamePrefix).AddArgument($CandidateEmailDomain).AddArgument($staticScreeningFields)
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
    $triggeredInBatch = @($batchResults | Where-Object FlowSuccess).Count
    Write-Host "  Created: $createdInBatch/$($batch.Count), Flow triggered: $triggeredInBatch/$($batch.Count)"
}

$allResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$totalCreated = @($allResults | Where-Object CreateSuccess).Count
$totalTriggered = @($allResults | Where-Object FlowSuccess).Count
$totalFailed = $allResults.Count - $totalTriggered

Write-Host ''
Write-Host "Done. Screenings created: $totalCreated/$($allResults.Count). Flows triggered: $totalTriggered/$($allResults.Count)."
if ($totalFailed -gt 0) {
    Write-Warning "$totalFailed record(s) had a create or flow-trigger failure. See $csvPath for per-record details."
}
else {
    Write-Host "Results saved to: $csvPath"
}
