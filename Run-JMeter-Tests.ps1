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

    [string]$ResultsDir = "$PSScriptRoot\results",
    [int]$Threads = 10,
    [int]$RampUp = 30,
    [int]$Loops = 1,
    # Leave at 0 to auto-scale heap from the actual thread count (recommended).
    # Pass an explicit value to pin the heap size and disable auto-scaling.
    [int]$MinHeapMB = 0,
    [int]$MaxHeapMB = 0,
    [int]$HeapMBPerThread = 6,
    [int]$BaseHeapMB = 512
)

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
        ThreadCounts        = @($threadCounts | Select-Object -Unique)
        DetectedThreadCount = $detectedThreadCount
    }
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

# Detect the domain/thread-count actually baked into the .jmx being run, since
# different test plans hardcode different values which silently override -Threads
# or send requests to a different host than -EnvironmentUrl.
$planInfo = Get-JMeterTestPlanEnvironmentInfo -TestPlanPath $TestPlanPath

if ($planInfo.Domains.Count -gt 0) {
    $envHost = ([System.Uri]$EnvironmentUrl).Host
    $mismatchedDomains = @($planInfo.Domains | Where-Object { $_ -ne $envHost })
    if ($mismatchedDomains.Count -gt 0) {
        Write-Warning "Test plan '$TestPlanPath' hardcodes HTTPSampler.domain '$($mismatchedDomains -join ', ')', which does NOT match -EnvironmentUrl host '$envHost'. JMeter will send its actual HTTP requests to the domain baked into the .jmx (NOT to -EnvironmentUrl)."
    }
}

if ($planInfo.ThreadCounts.Count -gt 1) {
    Write-Warning "Test plan '$TestPlanPath' hardcodes multiple different ThreadGroup.num_threads values: $($planInfo.ThreadCounts -join ', '). -Threads only affects thread groups that reference `${__P(thread_count)}`."
}

$effectiveThreadCount = $Threads
if ($planInfo.DetectedThreadCount) {
    $effectiveThreadCount = $planInfo.DetectedThreadCount
    if ($effectiveThreadCount -ne $Threads) {
        Write-Warning "Test plan '$TestPlanPath' hardcodes ThreadGroup.num_threads=$($planInfo.DetectedThreadCount), which does NOT match -Threads $Threads. JMeter will actually run with $($planInfo.DetectedThreadCount) threads."
    }
}

# Authentication

$scope = $EnvironmentUrl.TrimEnd('/') + "/.default"

# Request device code
$deviceResponse = Invoke-RestMethod `
    -Method POST `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } `
    -ContentType "application/x-www-form-urlencoded"

Write-Host $deviceResponse.message
Write-Host ""

# Keep contacting the server until login is done
$tokenBody = @{
    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
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
            -ContentType "application/x-www-form-urlencoded" `
            -ErrorAction Stop
        $accessToken = $tokenResponse.access_token
    }
    catch {
        $err = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue).error
        if ($err -eq "authorization_pending") { continue }
        elseif ($err -eq "slow_down") { $interval += 5; continue }
        elseif ($err -eq "authorization_declined") { Write-Error "Login was declined."; exit 1 }
        elseif ($err -eq "expired_token") { Write-Error "Device code expired."; exit 1 }
        else { Write-Error $_; exit 1 }
    }
}

Write-Host "Authenticated successfully."

# Running JMeter Test Plan

$hostname = ([System.Uri]$EnvironmentUrl).Host
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$runDir = Join-Path $ResultsDir $timestamp
$jtlFile = Join-Path $runDir "results.jtl"
$reportDir = Join-Path $runDir "report"

New-Item -ItemType Directory -Path $runDir -Force | Out-Null

# Auto-scale heap from actual thread concurrency unless the caller pins an explicit
# size: each JMeter thread holds its own connection/response/assertion state, so a
# fixed heap silently OOMs as thread count grows regardless of total request count.
if (-not ($PSBoundParameters.ContainsKey('MinHeapMB') -or $PSBoundParameters.ContainsKey('MaxHeapMB'))) {
    $computedMaxHeapMB = $BaseHeapMB + ($effectiveThreadCount * $HeapMBPerThread)

    $totalPhysicalMB = [math]::Floor((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB)
    $heapCeilingMB = [math]::Floor($totalPhysicalMB * 0.7)

    if ($computedMaxHeapMB -gt $heapCeilingMB) {
        Write-Warning "Auto-scaled heap for $effectiveThreadCount threads would be ${computedMaxHeapMB}MB, exceeding 70% of this machine's ${totalPhysicalMB}MB RAM. Capping -Xmx at ${heapCeilingMB}MB; reduce -Threads or run from a machine with more RAM if this is insufficient."
        $computedMaxHeapMB = $heapCeilingMB
    }

    $MaxHeapMB = [math]::Max($computedMaxHeapMB, 1024)
    $MinHeapMB = $MaxHeapMB
    Write-Host "Auto-scaled JVM heap for $effectiveThreadCount threads: -Xms${MinHeapMB}m -Xmx${MaxHeapMB}m ($HeapMBPerThread MB/thread + ${BaseHeapMB}MB base). Pass -MaxHeapMB to override."
}

# Configure JVM heap and GC options for large runs.
$heapString = "-Xms${MinHeapMB}m -Xmx${MaxHeapMB}m"
$jvmArgs = "-XX:+UseG1GC -XX:MaxGCPauseMillis=200 -XX:+HeapDumpOnOutOfMemoryError -XX:HeapDumpPath=$runDir\heapdump.hprof"
# Always overwrite: these are persistent per-terminal env vars, so a stale value
# from an earlier run in this same session would otherwise silently make
# -MinHeapMB/-MaxHeapMB have no effect.
$env:HEAP = $heapString
$env:JVM_ARGS = $jvmArgs

Write-Host "JMeter HEAP set to: $env:HEAP"
Write-Host "JVM_ARGS set to: $env:JVM_ARGS"

& $JMeterPath `
    -n `
    -t $TestPlanPath `
    -l $jtlFile `
    "-Jauth_token=Bearer $accessToken" `
    "-Jbase_url=$hostname" `
    "-Jthread_count=$Threads" `
    "-Jramp_up=$RampUp" `
    "-Jloop_count=$Loops" `
    "-Jjmeter.save.saveservice.output_format=csv" `
    "-Jjmeter.save.saveservice.response_data=false" `
    "-Jjmeter.save.saveservice.samplerData=false" `
    "-Jjmeter.save.saveservice.response_headers=false" `
    "-Jjmeter.save.saveservice.request_headers=false"

if ($LASTEXITCODE -eq 0) {
    Write-Host "Test run complete. Results saved to: $jtlFile"

    # Report generation is a separate JVM run from the load test on purpose: dashboard
    # synthesis needs memory proportional to total sample count, not thread concurrency,
    # so reusing the thread-count-scaled load-test heap here previously starved it and
    # surfaced as "Error generating the report: java.lang.NullPointerException".
    $jtlSizeMB = [math]::Ceiling((Get-Item $jtlFile).Length / 1MB)
    $reportHeapMB = [math]::Max(1024, $jtlSizeMB * 20)
    $totalPhysicalMB = [math]::Floor((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB)
    $reportHeapCeilingMB = [math]::Floor($totalPhysicalMB * 0.7)
    if ($reportHeapMB -gt $reportHeapCeilingMB) {
        $reportHeapMB = $reportHeapCeilingMB
    }
    $env:HEAP = "-Xms${reportHeapMB}m -Xmx${reportHeapMB}m"
    $env:JVM_ARGS = "-XX:+UseG1GC"

    Write-Host "Generating HTML report (heap: -Xmx${reportHeapMB}m, sized from ${jtlSizeMB}MB results file)..."
    & $JMeterPath -g $jtlFile -o $reportDir

    $reportIndex = Join-Path $reportDir "index.html"
    if ($LASTEXITCODE -eq 0 -and (Test-Path $reportIndex)) {
        Write-Host "Report saved to: $reportDir"
        Start-Process $reportIndex
    }
    else {
        Write-Warning "Report generation failed (exit code $LASTEXITCODE). Raw results are still available at $jtlFile; retry the dashboard alone with: & '$JMeterPath' -g '$jtlFile' -o '$reportDir'"
    }
}
else {
    Write-Error "JMeter exited with code $LASTEXITCODE. Check $runDir for details."
}
