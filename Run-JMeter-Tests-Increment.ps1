param(
    [Parameter(Mandatory)]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$JMeterPath,

    [Parameter(Mandatory)]
    [string]$TestPlanPath,

    [string]$RecordIdsPath,

    [string]$QueryUrl,
    [string]$QueryRequestLine,
    [string]$QueryEntityPath = 'cr9da_fsn_screeningses',
    [string]$QueryString,
    # Different .jmx files can target different Dataverse environments/API versions
    # (e.g. testplan.jmx uses /api/data/v9.2/..., updateSummaryReport_STAGE.jmx uses
    # /api/data/v9.0/... against a different hardcoded host). By default this is
    # auto-detected from -TestPlanPath's HTTPSampler.path values. Only pass this
    # explicitly to override auto-detection.
    [string]$ApiVersion = '9.0',
    [string]$FirstName,
    [string]$LastName,
    [string]$AdjudicationStatus,
    [string]$ScreeningStatus,
    [string]$AssignedToEmailFilter,
    [string]$CandidateEmail,
    [string]$BusinessLine,
    [string]$AdditionalFilter,
    [switch]$UseContains,
    # Defaults to the effective thread count detected from -TestPlanPath (or -Threads
    # if the .jmx doesn't hardcode a literal ThreadGroup.num_threads), capped at 2000.
    # Pass this explicitly to override. Always clamped to a maximum of 2000.
    [int]$Top = 2000,
    [int]$MaxRows = 20000,

    [ValidateSet('PerThread', 'Shared')]
    [string]$DistributionMode = 'PerThread',

    [string]$ResultsDir = "$PSScriptRoot\results",
    [int]$Threads = 10,
    [int]$RampUp = 30,
    [int]$Loops = 1,
    [int]$MinHeapMB = 1024,
    [int]$MaxHeapMB = 4096,

    [string]$RecordIdVariableName = 'record_id',

    [string]$AssignedToDisplayName,
    [string]$AssignedToEmail,
    [string]$AssignedToDate
)

function Convert-ToJMeterPath {
    param([Parameter(Mandatory)][string]$Path)
    return $Path -replace '\\', '/'
}

function Read-RecordIds {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$VariableName
    )

    if (-not (Test-Path $Path)) {
        throw "Record IDs file not found at '$Path'."
    }

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $values = @()

    if ($ext -eq '.csv') {
        $rows = Import-Csv -Path $Path
        if (-not $rows -or $rows.Count -eq 0) {
            throw "CSV '$Path' is empty."
        }

        $candidateColumns = @(
            $VariableName,
            'record_id',
            'cr9da_fsn_screeningsid',
            'id'
        )

        $column = $null
        foreach ($c in $candidateColumns) {
            if ($rows[0].PSObject.Properties.Name -contains $c) {
                $column = $c
                break
            }
        }

        if (-not $column) {
            $firstProp = $rows[0].PSObject.Properties | Select-Object -First 1
            if (-not $firstProp) {
                throw "CSV '$Path' has no readable columns."
            }
            $column = $firstProp.Name
        }

        $values = $rows | ForEach-Object { $_.$column }
        Write-Host "Using CSV column '$column' for record IDs."
    }
    else {
        $values = Get-Content -Path $Path | ForEach-Object { $_.Trim() }
    }

    $clean = @(
        $values |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim().Trim('{}').ToLowerInvariant() } |
        Select-Object -Unique
    )

    if ($clean.Count -eq 0) {
        throw "No record IDs were found in '$Path'."
    }

    $guidRegex = '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
    $invalid = @($clean | Where-Object { $_ -notmatch $guidRegex })
    if ($invalid.Count -gt 0) {
        $sample = $invalid | Select-Object -First 5
        throw "Found non-GUID record IDs: $($sample -join ', ')"
    }

    return $clean
}

function ConvertTo-ODataEscapedString {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace("'", "''")
}

function Get-JMeterTestPlanEnvironmentInfo {
    param([Parameter(Mandatory)][string]$TestPlanPath)

    $xmlText = Get-Content -Path $TestPlanPath -Raw

    $domains = [System.Collections.Generic.List[string]]::new()
    foreach ($m in [regex]::Matches($xmlText, '<stringProp name="HTTPSampler\.domain">([^<]*)</stringProp>')) {
        $val = $m.Groups[1].Value.Trim()
        if (-not [string]::IsNullOrWhiteSpace($val) -and $val -notmatch '\$\{') {
            $domains.Add($val)
        }
    }

    $versions = [System.Collections.Generic.List[string]]::new()
    foreach ($m in [regex]::Matches($xmlText, '/api/data/v([0-9]+\.[0-9]+)/')) {
        $versions.Add($m.Groups[1].Value)
    }

    $detectedVersion = $null
    if ($versions.Count -gt 0) {
        $detectedVersion = ($versions | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
    }

    $threadCounts = [System.Collections.Generic.List[int]]::new()
    foreach ($m in [regex]::Matches($xmlText, '<stringProp name="ThreadGroup\.num_threads">([^<]*)</stringProp>')) {
        $val = $m.Groups[1].Value.Trim()
        if ($val -notmatch '\$\{' -and $val -match '^[0-9]+$') {
            $threadCounts.Add([int]$val)
        }
    }

    $detectedThreadCount = $null
    if ($threadCounts.Count -gt 0) {
        $detectedThreadCount = ($threadCounts | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name -as [int]
    }

    return [pscustomobject]@{
        Domains             = @($domains | Select-Object -Unique)
        ApiVersions         = @($versions | Select-Object -Unique)
        DetectedApiVersion  = $detectedVersion
        ThreadCounts        = @($threadCounts | Select-Object -Unique)
        DetectedThreadCount = $detectedThreadCount
    }
}

function New-FilterClause {
    param(
        [Parameter(Mandatory)][string]$Field,
        [Parameter(Mandatory)][string]$Value,
        [Parameter(Mandatory)][bool]$ContainsMode
    )

    $escaped = ConvertTo-ODataEscapedString -Value $Value
    if ($ContainsMode) {
        return "contains($Field,'$escaped')"
    }

    return "$Field eq '$escaped'"
}

function Get-FriendlyHttpError {
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [string]$RequestUrl
    )

    $statusCode = $null
    $statusText = $null
    if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Response) {
        try {
            $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode
            $statusText = [string]$ErrorRecord.Exception.Response.StatusCode
        }
        catch {
            $statusCode = $null
            $statusText = $null
        }
    }

    $detail = $ErrorRecord.ErrorDetails.Message
    if ([string]::IsNullOrWhiteSpace($detail)) {
        $detail = $ErrorRecord.Exception.Message
    }

    if (-not [string]::IsNullOrWhiteSpace($detail) -and $detail -match '<html|<!doctype html|Server Error in ''/'' Application') {
        $detail = 'Server returned an HTML error page instead of Dataverse JSON. This usually means the URL is not a Dataverse Web API endpoint.'
    }

    if ($RequestUrl) {
        if ($statusCode) {
            return "Dataverse query failed with HTTP $statusCode ($statusText). URL: $RequestUrl. Details: $detail"
        }
        return "Dataverse query failed. URL: $RequestUrl. Details: $detail"
    }

    if ($statusCode) {
        return "Dataverse query failed with HTTP $statusCode ($statusText). Details: $detail"
    }

    return "Dataverse query failed. Details: $detail"
}

function Convert-RequestLineToDataverseUrl {
    param(
        [Parameter(Mandatory)][string]$EnvironmentUrl,
        [Parameter(Mandatory)][string]$RequestLine,
        [string]$ApiVersion = '9.2'
    )

    $trimmed = $RequestLine.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        throw '-QueryRequestLine cannot be empty.'
    }

    $pathAndQuery = $trimmed
    if ($trimmed -match '^GET\s+(.+?)\s+HTTP/\d(?:\.\d)?\s*$') {
        $pathAndQuery = $Matches[1]
    }

    if ($pathAndQuery -match '^https?://') {
        return $pathAndQuery
    }

    $envBase = $EnvironmentUrl.TrimEnd('/')

    if ($pathAndQuery.StartsWith('/api/data/', [System.StringComparison]::OrdinalIgnoreCase)) {
        return "$envBase$pathAndQuery"
    }

    if ($pathAndQuery.StartsWith('/')) {
        return "$envBase$pathAndQuery"
    }

    return "$envBase/api/data/v$ApiVersion/$pathAndQuery"
}

function Get-RecordIdsFromQuery {
    param(
        [Parameter(Mandatory)][string]$EnvironmentUrl,
        [Parameter(Mandatory)][string]$AccessToken,
        [string]$QueryUrl,
        [string]$QueryRequestLine,
        [Parameter(Mandatory)][string]$QueryEntityPath,
        [string]$QueryString,
        [string]$FirstName,
        [string]$LastName,
        [string]$AdjudicationStatus,
        [string]$ScreeningStatus,
        [string]$AssignedToEmailFilter,
        [string]$CandidateEmail,
        [string]$BusinessLine,
        [string]$AdditionalFilter,
        [bool]$UseContains,
        [int]$Top,
        [int]$MaxRows,
        [string]$ApiVersion = '9.2'
    )

    if (-not [string]::IsNullOrWhiteSpace($QueryUrl)) {
        if (-not [System.Uri]::IsWellFormedUriString($QueryUrl, [System.UriKind]::Absolute)) {
            throw "-QueryUrl is not a valid absolute URL: '$QueryUrl'"
        }
        if ($QueryUrl -notmatch '/api/data/v[0-9.]+/') {
            throw "-QueryUrl must target the Dataverse Web API path (/api/data/vX.X/...). Received: '$QueryUrl'"
        }
        $resolvedUrl = $QueryUrl
    }
    elseif (-not [string]::IsNullOrWhiteSpace($QueryRequestLine)) {
        $resolvedUrl = Convert-RequestLineToDataverseUrl -EnvironmentUrl $EnvironmentUrl -RequestLine $QueryRequestLine -ApiVersion $ApiVersion
        if ($resolvedUrl -notmatch '/api/data/v[0-9.]+/') {
            throw "-QueryRequestLine must resolve to a Dataverse Web API path (/api/data/vX.X/...). Resolved: '$resolvedUrl'"
        }
    }
    else {
        $entityPath = $QueryEntityPath
        if ([string]::IsNullOrWhiteSpace($entityPath) -or $entityPath.TrimStart().StartsWith('$')) {
            $entityPath = 'cr9da_fsn_screeningses'
        }
        $entityPath = $entityPath.Trim().Trim('/')

        if ($Top -le 0) {
            throw '-Top must be greater than 0.'
        }
        if ($MaxRows -le 0) {
            throw '-MaxRows must be greater than 0.'
        }

        $queryPart = $QueryString
        if (-not [string]::IsNullOrWhiteSpace($queryPart)) {
            $queryPart = $queryPart.TrimStart('?')
        }
        else {
            $filterParts = [System.Collections.Generic.List[string]]::new()

            if ($FirstName) { $filterParts.Add((New-FilterClause -Field 'cr9da_firstname' -Value $FirstName -ContainsMode $UseContains)) }
            if ($LastName) { $filterParts.Add((New-FilterClause -Field 'cr9da_lastname' -Value $LastName -ContainsMode $UseContains)) }
            if ($AdjudicationStatus) { $filterParts.Add((New-FilterClause -Field 'cr9da_adjudicationphase' -Value $AdjudicationStatus -ContainsMode $UseContains)) }
            if ($ScreeningStatus) { $filterParts.Add((New-FilterClause -Field 'cr9da_screeningstatus' -Value $ScreeningStatus -ContainsMode $UseContains)) }
            if ($AssignedToEmailFilter) { $filterParts.Add((New-FilterClause -Field 'cr9da_assignedtoemail' -Value $AssignedToEmailFilter -ContainsMode $UseContains)) }
            if ($CandidateEmail) { $filterParts.Add((New-FilterClause -Field 'cr9da_candidateemail' -Value $CandidateEmail -ContainsMode $UseContains)) }
            if ($BusinessLine) { $filterParts.Add((New-FilterClause -Field 'cr9da_businessline' -Value $BusinessLine -ContainsMode $UseContains)) }
            if ($AdditionalFilter) { $filterParts.Add("($AdditionalFilter)") }

            if ($filterParts.Count -eq 0) {
                throw 'When -RecordIdsPath is not provided, pass at least one filter, -QueryString, or -QueryUrl.'
            }

            $selectColumns = @('cr9da_fsn_screeningsid', 'createdon')
            $parts = [System.Collections.Generic.List[string]]::new()
            $parts.Add('$select=' + [System.Uri]::EscapeDataString(($selectColumns -join ',')))
            $parts.Add('$orderby=' + [System.Uri]::EscapeDataString('createdon desc'))
            $parts.Add('$top=' + $Top)
            $parts.Add('$filter=' + [System.Uri]::EscapeDataString(($filterParts -join ' and ')))
            $queryPart = $parts -join '&'
        }

        if ([string]::IsNullOrWhiteSpace($queryPart)) {
            throw 'No query parameters were built. Pass a filter, -QueryString, or -QueryUrl.'
        }

        $resolvedUrl = $EnvironmentUrl.TrimEnd('/') + '/api/data/v' + $ApiVersion + '/' + $entityPath + '?' + $queryPart
    }

    Write-Host "Querying records from: $resolvedUrl"

    $headers = @{ Authorization = "Bearer $AccessToken"; Accept = 'application/json' }
    $rows = [System.Collections.Generic.List[object]]::new()
    $nextUrl = $resolvedUrl

    while (-not [string]::IsNullOrWhiteSpace($nextUrl)) {
        try {
            $response = Invoke-RestMethod -Method GET -Uri $nextUrl -Headers $headers -ErrorAction Stop
        }
        catch {
            throw (Get-FriendlyHttpError -ErrorRecord $_ -RequestUrl $nextUrl)
        }

        if ($response.value) {
            foreach ($row in $response.value) {
                $rows.Add($row)
                if ($rows.Count -ge $MaxRows) {
                    break
                }
            }
        }

        if ($rows.Count -ge $MaxRows) {
            Write-Warning "Reached MaxRows=$MaxRows. Stopping pagination early."
            break
        }

        $nextUrl = $null
        if ($response.PSObject.Properties.Name -contains '@odata.nextLink') {
            $nextUrl = $response.'@odata.nextLink'
        }
    }

    $ids = @(
        $rows |
        ForEach-Object { $_.cr9da_fsn_screeningsid } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.ToString().Trim().Trim('{}').ToLowerInvariant() } |
        Select-Object -Unique
    )

    if ($ids.Count -eq 0) {
        throw 'Query returned zero screening GUIDs.'
    }

    Write-Host "Query returned $($rows.Count) rows and $($ids.Count) unique GUIDs."
    return $ids
}

# Checks before running script
if (-not (Test-Path $JMeterPath)) {
    Write-Error "JMeter not found at '$JMeterPath'. Pass -JMeterPath to your jmeter.bat location."
    exit 1
}

if (-not (Test-Path $TestPlanPath)) {
    Write-Error "Test plan not found at '$TestPlanPath'."
    exit 1
}

if (-not ([System.Uri]::IsWellFormedUriString($EnvironmentUrl, [System.UriKind]::Absolute))) {
    Write-Error "EnvironmentUrl '$EnvironmentUrl' is not a valid URL."
    exit 1
}

if ($EnvironmentUrl -notmatch 'https://[^/]+\.crm[0-9]*\.dynamics\.com/?$') {
    Write-Warning "EnvironmentUrl does not look like a standard Dataverse org URL (for example: https://org.crm9.dynamics.com). Current value: '$EnvironmentUrl'"
}

# Detect the Dataverse API version (and any hardcoded host) actually used by the
# .jmx being run, since different test plans target different environments/versions.
$planInfo = Get-JMeterTestPlanEnvironmentInfo -TestPlanPath $TestPlanPath

if ($planInfo.ApiVersions.Count -gt 1) {
    Write-Warning "Test plan '$TestPlanPath' references multiple Dataverse API versions: $($planInfo.ApiVersions -join ', '). Using the most common one: $($planInfo.DetectedApiVersion)."
}

if (-not $PSBoundParameters.ContainsKey('ApiVersion') -and $planInfo.DetectedApiVersion) {
    if ($ApiVersion -ne $planInfo.DetectedApiVersion) {
        Write-Host "Auto-detected Dataverse API version 'v$($planInfo.DetectedApiVersion)' from '$TestPlanPath'."
    }
    $ApiVersion = $planInfo.DetectedApiVersion
}
elseif (-not $planInfo.DetectedApiVersion) {
    Write-Warning "Could not detect an API version from '$TestPlanPath'. Falling back to -ApiVersion '$ApiVersion'."
}

if ($planInfo.Domains.Count -gt 0) {
    $envHost = ([System.Uri]$EnvironmentUrl).Host
    $mismatchedDomains = @($planInfo.Domains | Where-Object { $_ -ne $envHost })
    if ($mismatchedDomains.Count -gt 0) {
        Write-Warning "Test plan '$TestPlanPath' hardcodes HTTPSampler.domain '$($mismatchedDomains -join ', ')', which does NOT match -EnvironmentUrl host '$envHost'. JMeter will send its actual HTTP requests to the domain baked into the .jmx (NOT to -EnvironmentUrl), while this script queries Dataverse for record GUIDs using -EnvironmentUrl. Point -EnvironmentUrl at '$($mismatchedDomains[0])' (or update the .jmx) so the GUIDs you query match the environment JMeter actually hits."
    }
}

# Determine the effective thread count actually used by this run: if the .jmx
# hardcodes a literal ThreadGroup.num_threads (instead of referencing
# ${__P(thread_count)}), -Threads has NO effect on JMeter and the .jmx value wins.
if ($planInfo.ThreadCounts.Count -gt 1) {
    Write-Warning "Test plan '$TestPlanPath' hardcodes multiple different ThreadGroup.num_threads values: $($planInfo.ThreadCounts -join ', '). -Threads only affects thread groups that reference `${__P(thread_count)}`."
}

$effectiveThreadCount = $Threads
if ($planInfo.DetectedThreadCount) {
    $effectiveThreadCount = $planInfo.DetectedThreadCount
    if ($effectiveThreadCount -ne $Threads) {
        Write-Warning "Test plan '$TestPlanPath' hardcodes ThreadGroup.num_threads=$($planInfo.DetectedThreadCount), which does NOT match -Threads $Threads. JMeter will actually run with $($planInfo.DetectedThreadCount) threads (the value baked into the .jmx). Update the .jmx to use `${__P(thread_count,...)}` (or pass -Threads $($planInfo.DetectedThreadCount)) to keep these consistent."
    }
}

# Default -Top to the effective thread count, capped at 2000, unless explicitly passed.
if (-not $PSBoundParameters.ContainsKey('Top')) {
    $Top = $effectiveThreadCount
}
if ($Top -gt 2000) {
    Write-Warning "-Top $Top exceeds the maximum of 2000; clamping to 2000."
    $Top = 2000
}
Write-Host "Using -Top $Top for the Dataverse GUID query (effective thread count: $effectiveThreadCount)."

# Build per-run output locations
$hostname = ([System.Uri]$EnvironmentUrl).Host
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$runDir = Join-Path $ResultsDir $timestamp
$jtlFile = Join-Path $runDir 'results.jtl'
$reportDir = Join-Path $runDir 'report'
$recordsDir = Join-Path $runDir 'records'

New-Item -ItemType Directory -Path $runDir -Force | Out-Null
New-Item -ItemType Directory -Path $recordsDir -Force | Out-Null

# Authentication
$scope = $EnvironmentUrl.TrimEnd('/') + '/.default'

$deviceResponse = Invoke-RestMethod `
    -Method POST `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } `
    -ContentType 'application/x-www-form-urlencoded'

Write-Host $deviceResponse.message
Write-Host ''

$tokenBody = @{
    grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
    client_id   = $ClientId
    device_code = $deviceResponse.device_code
}

$accessToken = $null
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

# Build record ID list from local file or query mode.
try {
    if (-not [string]::IsNullOrWhiteSpace($RecordIdsPath)) {
        $recordIds = Read-RecordIds -Path $RecordIdsPath -VariableName $RecordIdVariableName
        Write-Host "Loaded $($recordIds.Count) unique record IDs from '$RecordIdsPath'."
    }
    else {
        $recordIds = Get-RecordIdsFromQuery `
            -EnvironmentUrl $EnvironmentUrl `
            -AccessToken $accessToken `
            -QueryUrl $QueryUrl `
            -QueryRequestLine $QueryRequestLine `
            -QueryEntityPath $QueryEntityPath `
            -QueryString $QueryString `
            -FirstName $FirstName `
            -LastName $LastName `
            -AdjudicationStatus $AdjudicationStatus `
            -ScreeningStatus $ScreeningStatus `
            -AssignedToEmailFilter $AssignedToEmailFilter `
            -CandidateEmail $CandidateEmail `
            -BusinessLine $BusinessLine `
            -AdditionalFilter $AdditionalFilter `
            -UseContains $UseContains.IsPresent `
            -Top $Top `
            -MaxRows $MaxRows `
            -ApiVersion $ApiVersion
    }
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}

if ($DistributionMode -eq 'PerThread' -and $recordIds.Count -lt $Threads) {
    Write-Warning "You provided fewer IDs ($($recordIds.Count)) than threads ($Threads). Some thread CSV files will be empty and may stop quickly depending on your CSVDataSet settings."
}

# Create CSV files used by JMeter.
$recordsHeader = $RecordIdVariableName
$recordProps = @()

if ($DistributionMode -eq 'Shared') {
    $sharedCsv = Join-Path $recordsDir 'record_ids.csv'
    $content = @($recordsHeader) + $recordIds
    Set-Content -Path $sharedCsv -Value $content -Encoding UTF8

    $recordProps += "-Jrecord_ids_file=$(Convert-ToJMeterPath -Path $sharedCsv)"
    $recordProps += "-Jrecord_ids_mode=shared"

    Write-Host "Created shared record CSV: $sharedCsv"
}
else {
    $bucketed = @{}
    for ($i = 1; $i -le $Threads; $i++) {
        $bucketed[$i] = [System.Collections.Generic.List[string]]::new()
    }

    for ($idx = 0; $idx -lt $recordIds.Count; $idx++) {
        $threadNum = ($idx % $Threads) + 1
        $bucketed[$threadNum].Add($recordIds[$idx])
    }

    for ($threadNum = 1; $threadNum -le $Threads; $threadNum++) {
        $threadCsv = Join-Path $recordsDir "record_ids_thread_${threadNum}.csv"
        $content = @($recordsHeader) + $bucketed[$threadNum]
        Set-Content -Path $threadCsv -Value $content -Encoding UTF8
    }

    $pattern = Join-Path $recordsDir 'record_ids_thread_${__threadNum}.csv'
    $recordProps += "-Jrecord_ids_file_pattern=$(Convert-ToJMeterPath -Path $pattern)"
    $recordProps += "-Jrecord_ids_mode=per_thread"

    Write-Host "Created $Threads per-thread record CSV files in: $recordsDir"
}

$recordProps += "-Jrecord_id_var_name=$RecordIdVariableName"

if ($AssignedToDisplayName) { $recordProps += "-Jassigned_to_display_name=$AssignedToDisplayName" }
if ($AssignedToEmail) { $recordProps += "-Jassigned_to_email=$AssignedToEmail" }
if ($AssignedToDate) { $recordProps += "-Jassigned_to_date=$AssignedToDate" }

# Configure JVM heap and GC options for large runs (can be overridden by environment)
$heapString = "-Xms${MinHeapMB}m -Xmx${MaxHeapMB}m"
$jvmArgs = "-XX:+UseG1GC -XX:MaxGCPauseMillis=200 -XX:+HeapDumpOnOutOfMemoryError -XX:HeapDumpPath=$runDir\heapdump.hprof"
if (-not $env:HEAP) { $env:HEAP = $heapString }
if (-not $env:JVM_ARGS) { $env:JVM_ARGS = $jvmArgs }

Write-Host "JMeter HEAP set to: $env:HEAP"
Write-Host "JVM_ARGS set to: $env:JVM_ARGS"

$jmeterArgs = @(
    '-n',
    '-t', $TestPlanPath,
    '-l', $jtlFile,
    '-e',
    '-o', $reportDir,
    "-Jauth_token=Bearer $accessToken",
    "-Jbase_url=$hostname",
    "-Jthread_count=$Threads",
    "-Jramp_up=$RampUp",
    "-Jloop_count=$Loops",
    '-Jjmeter.save.saveservice.output_format=csv',
    '-Jjmeter.save.saveservice.response_data=false',
    '-Jjmeter.save.saveservice.samplerData=false',
    '-Jjmeter.save.saveservice.response_headers=false',
    '-Jjmeter.save.saveservice.request_headers=false'
)

$jmeterArgs += $recordProps

Write-Host 'Launching JMeter with record-increment properties:'
$recordProps | ForEach-Object { Write-Host "  $_" }

& $JMeterPath @jmeterArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "Report saved to: $reportDir"
    Start-Process (Join-Path $reportDir 'index.html')
}
else {
    Write-Error "JMeter exited with code $LASTEXITCODE. Check $runDir for details."
}
