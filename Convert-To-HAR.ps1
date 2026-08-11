param(
  [Parameter(Mandatory)][string]$HarFile,
  [string]$OutputJmx = '',
  [string]$UrlFilter = '',
  [string]$ApiBase = '/api/data/v9.0',
  [int]$Threads = 10,
  [int]$RampUp = -1,
  [int]$Loops = 1,
  [string]$RecordIdGuid = ''
)

function ConvertTo-XmlText([string]$s) {
  [System.Security.SecurityElement]::Escape($s)
}

function Get-BatchSubRequests([string]$Body, [string]$MimeType, [string]$ApiBase) {
  if ($MimeType -notmatch 'boundary=([^\s;]+)') { return @() }
  $boundary = $Matches[1].Trim('"')
  $results = [System.Collections.Generic.List[object]]::new()

  foreach ($part in ($Body -split "--$([regex]::Escape($boundary))")) {
    $trimmed = $part.Trim("`r`n ")
    if (-not $trimmed -or $trimmed -eq '--') { continue }

    $lines = $trimmed -split "`r?`n"

    $innerStart = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
      if ($lines[$i].Trim() -eq '') { $innerStart = $i + 1; break }
    }
    if ($innerStart -lt 0 -or $innerStart -ge $lines.Count) { continue }

    $tokens = $lines[$innerStart].Trim() -split '\s+'
    if ($tokens.Count -lt 2) { continue }

    $innerMethod = $tokens[0].ToUpper()
    $innerPath = $tokens[1]
    if ($innerPath -like 'http*') {
      $innerPath = ([System.Uri]$innerPath).PathAndQuery
    }
    elseif (-not $innerPath.StartsWith('/')) {
      $innerPath = "$ApiBase/$innerPath"
    }

    $innerHeaders = @{}
    $bodyStart = -1
    for ($i = $innerStart + 1; $i -lt $lines.Count; $i++) {
      if ($lines[$i].Trim() -eq '') { $bodyStart = $i + 1; break }
      $colonIdx = $lines[$i].IndexOf(':')
      if ($colonIdx -gt 0) {
        $innerHeaders[$lines[$i].Substring(0, $colonIdx).Trim()] = $lines[$i].Substring($colonIdx + 1).Trim()
      }
    }

    $innerBody = ''
    if ($bodyStart -gt 0 -and $bodyStart -lt $lines.Count) {
      $innerBody = ($lines[$bodyStart..($lines.Count - 1)] -join "`n").Trim()
    }

    $results.Add([PSCustomObject]@{
        Method  = $innerMethod
        Path    = $innerPath
        Headers = $innerHeaders
        Body    = $innerBody
      })
  }

  return $results
}

function Get-SamplerLabel([string]$Method, [string]$Path) {
  $rel = $Path -replace '^/api/data/v[\d.]+/', ''
  $selectIdx = $rel.IndexOf('select', [System.StringComparison]::OrdinalIgnoreCase)
  if ($selectIdx -gt 0) { "$Method $($rel.Substring(0, $selectIdx))" }
  else { "$Method $rel" }
}

if ($RampUp -lt 0) { $RampUp = $Threads }

if (-not (Test-Path $HarFile)) { Write-Error "HAR file not found: $HarFile"; exit 1 }
if (-not $OutputJmx) { $OutputJmx = [System.IO.Path]::ChangeExtension($HarFile, '.jmx') }

# Build property references (without backtick escaping in output)
$authTokenRef = '$' + '{__P(auth_token)}'
$baseUrlRef = '$' + '{__P(base_url)}'
$threadCountRef = '$' + "{__P(thread_count,$Threads)}"
$rampUpRef = '$' + "{__P(ramp_up,$RampUp)}"
# Single shared file (not one-per-thread): every thread reads/recycles the same
# CSV, so the jmx never hard-requires a specific thread count's worth of files to
# exist - it works standalone (default file, one row) or fed a bulk list via
# -Jrecord_ids_file from Run-JMeter-Tests-Increment.ps1.
$recordIdsFileRef = '$' + '{__P(record_ids_file,record_ids.csv)}'
$recordIdVarNameRef = '$' + '{__P(record_id_var_name,record_id)}'

$har = Get-Content $HarFile -Raw | ConvertFrom-Json
if (-not $har.log -or -not $har.log.entries) {
  Write-Error "HAR file has no log.entries - is it a valid HAR?"
  exit 1
}

Write-Host "$($har.log.entries.Count) total entries in HAR."

$allEntries = @($har.log.entries)
$entries = if ($UrlFilter) {
  @($allEntries | Where-Object { $_.request.url -like "*$UrlFilter*" })
}
else {
  @($allEntries)
}

if ($entries.Count -eq 0 -and $UrlFilter) {
  Write-Warning "No entries matched '$UrlFilter'. Showing all URLs and continuing with all entries:"
  $allEntries | ForEach-Object { Write-Host "  $($_.request.method) $($_.request.url)" }
  $entries = $allEntries
}

if ($entries.Count -eq 0) {
  Write-Error "No HAR entries were available to convert."
  exit 1
}

# Determine which GUID (if any) represents "the record this session captured",
# so it can be replaced with ${record_id} and driven per-thread by
# Run-JMeter-Tests-Increment.ps1's generated CSVs instead of staying hardcoded.
if (-not $RecordIdGuid) {
  # Exclude multipart $batch/changeset boundary GUIDs (always repeated, always
  # noise) via negative lookbehind - otherwise the boundary token gets
  # mistaken for a genuinely repeated record GUID.
  $guidPattern = '(?<!batch_)(?<!changeset_)[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
  $guidCounts = @{}
  foreach ($e in $entries) {
    $haystack = $e.request.url
    if ($e.request.postData -and $e.request.postData.text) { $haystack += "`n$($e.request.postData.text)" }
    foreach ($m in [regex]::Matches($haystack, $guidPattern)) {
      $g = $m.Value.ToLowerInvariant()
      if ($guidCounts.ContainsKey($g)) { $guidCounts[$g]++ } else { $guidCounts[$g] = 1 }
    }
  }
  $top = $guidCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
  if ($top -and $top.Value -gt 1) {
    $RecordIdGuid = $top.Key
    Write-Host "Auto-detected repeated record GUID '$RecordIdGuid' ($($top.Value) occurrences) - substituting with `${record_id}`. Pass -RecordIdGuid to override if this is wrong."
  }
}
else {
  $RecordIdGuid = $RecordIdGuid.ToLowerInvariant()
}

$baseHost = ([System.Uri]$entries[0].request.url).Host
$sb = [System.Text.StringBuilder]::new()
$nameCounts = @{}
$samplerCount = 0

foreach ($entry in $entries) {
  $req = $entry.request
  $uri = [System.Uri]$req.url
  $method = $req.method.ToUpper()

  $requests = @()
  if ($method -eq 'POST' -and ($uri.AbsolutePath -like '*/$batch' -or $uri.AbsolutePath -like '*/%24batch') -and $req.postData -and $req.postData.text) {
    $apiBase = $uri.AbsolutePath -replace '/(\$batch|%24batch)$', ''
    if (-not $apiBase) { $apiBase = $ApiBase }
    $requests = @(Get-BatchSubRequests -Body $req.postData.text -MimeType $req.postData.mimeType -ApiBase $apiBase)
    if ($requests.Count -eq 0) {
      Write-Warning "Could not parse batch body for $($uri.AbsolutePath) - adding as raw POST."
    }
  }

  if ($requests.Count -eq 0) {
    $rawPath = $req.url.Substring($req.url.IndexOf($baseHost) + $baseHost.Length)
    $ph = $req.headers | Where-Object { $_.name -ieq 'Prefer' } | Select-Object -First 1
    $requests = @([PSCustomObject]@{
        Method  = $method
        Path    = $rawPath
        Headers = @{
          Prefer         = if ($ph) { $ph.value } else { $null }
          'Content-Type' = if ($req.postData) { $req.postData.mimeType } else { $null }
        }
        Body    = if ($req.postData) { $req.postData.text } else { '' }
      })
  }

  foreach ($r in $requests) {
    $rawLabel = Get-SamplerLabel -Method $r.Method -Path $r.Path
    if ($nameCounts.ContainsKey($rawLabel)) {
      $nameCounts[$rawLabel]++
      $label = "$rawLabel ($($nameCounts[$rawLabel]))"
    }
    else {
      $nameCounts[$rawLabel] = 1
      $label = $rawLabel
    }

    if ($RecordIdGuid) {
      $r.Path = $r.Path -replace [regex]::Escape($RecordIdGuid), '${record_id}'
      if ($r.Body) { $r.Body = $r.Body -replace [regex]::Escape($RecordIdGuid), '${record_id}' }
    }

    $xmlLabel = ConvertTo-XmlText $label
    $xmlPath = ConvertTo-XmlText $r.Path
    $preferVal = if ($r.Headers['Prefer']) { $r.Headers['Prefer'] } else { 'odata.maxpagesize=2000,odata.include-annotations=*' }
    $xmlPrefer = ConvertTo-XmlText $preferVal

    $postBodyRaw = 'false'
    $bodyArgsXml = '            <collectionProp name="Arguments.arguments"/>'
    if ($r.Method -in @('POST', 'PATCH', 'PUT') -and $r.Body) {
      $postBodyRaw = 'true'
      $xmlBody = ConvertTo-XmlText $r.Body
      $bodyArgsXml = @"
            <collectionProp name="Arguments.arguments">
              <elementProp name="" elementType="HTTPArgument">
                <boolProp name="HTTPArgument.always_encode">false</boolProp>
                <stringProp name="Argument.value">$xmlBody</stringProp>
                <stringProp name="Argument.metadata">=</stringProp>
              </elementProp>
            </collectionProp>
"@
    }

    $perSamplerHeaders = @"
              <elementProp name="" elementType="Header">
                <stringProp name="Header.name">Prefer</stringProp>
                <stringProp name="Header.value">$xmlPrefer</stringProp>
              </elementProp>
"@
    if ($r.Method -in @('POST', 'PATCH', 'PUT') -and $r.Headers['Content-Type']) {
      $xmlCt = ConvertTo-XmlText $r.Headers['Content-Type']
      $perSamplerHeaders += @"
              <elementProp name="" elementType="Header">
                <stringProp name="Header.name">Content-Type</stringProp>
                <stringProp name="Header.value">$xmlCt</stringProp>
              </elementProp>
"@
    }

    $rMethod = $r.Method
    $null = $sb.Append(@"

        <HTTPSamplerProxy guiclass="HttpTestSampleGui" testclass="HTTPSamplerProxy" testname="$xmlLabel" enabled="true">
          <boolProp name="HTTPSampler.postBodyRaw">$postBodyRaw</boolProp>
          <elementProp name="HTTPsampler.Arguments" elementType="Arguments" guiclass="HTTPArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
$bodyArgsXml
          </elementProp>
          <stringProp name="HTTPSampler.path">$xmlPath</stringProp>
          <stringProp name="HTTPSampler.method">$rMethod</stringProp>
          <boolProp name="HTTPSampler.follow_redirects">true</boolProp>
          <boolProp name="HTTPSampler.use_keepalive">true</boolProp>
          <boolProp name="HTTPSampler.DO_MULTIPART_POST">false</boolProp>
        </HTTPSamplerProxy>
        <hashTree>
          <HeaderManager guiclass="HeaderPanel" testclass="HeaderManager" testname="Per-Request Headers" enabled="true">
            <collectionProp name="HeaderManager.headers">
$perSamplerHeaders
            </collectionProp>
          </HeaderManager>
          <hashTree/>
        </hashTree>
"@)
    $samplerCount++
  }
}

$harName = ConvertTo-XmlText (Split-Path $HarFile -Leaf)
$samplersXml = $sb.ToString()

$csvDataSetXml = ''
if ($RecordIdGuid) {
  $csvDataSetXml = @"
        <CSVDataSet guiclass="TestBeanGUI" testclass="CSVDataSet" testname="Record IDs" enabled="true">
          <stringProp name="delimiter">,</stringProp>
          <stringProp name="fileEncoding">UTF-8</stringProp>
          <stringProp name="filename">$recordIdsFileRef</stringProp>
          <boolProp name="ignoreFirstLine">true</boolProp>
          <boolProp name="quotedData">false</boolProp>
          <boolProp name="recycle">true</boolProp>
          <boolProp name="stopThread">false</boolProp>
          <stringProp name="variableNames">$recordIdVarNameRef</stringProp>
          <stringProp name="shareMode">shareMode.all</stringProp>
        </CSVDataSet>
        <hashTree/>
"@
}

$jmx = @"
<?xml version="1.0" encoding="UTF-8"?>
<jmeterTestPlan version="1.2" properties="5.0" jmeter="5.6.3">
  <hashTree>
    <TestPlan guiclass="TestPlanGui" testclass="TestPlan" testname="$harName" enabled="true">
      <boolProp name="TestPlan.functional_mode">false</boolProp>
      <boolProp name="TestPlan.tearDown_on_shutdown">true</boolProp>
      <boolProp name="TestPlan.serialize_threadgroups">false</boolProp>
      <elementProp name="TestPlan.user_defined_variables" elementType="Arguments" guiclass="ArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
        <collectionProp name="Arguments.arguments"/>
      </elementProp>
    </TestPlan>
    <hashTree>
      <ThreadGroup guiclass="ThreadGroupGui" testclass="ThreadGroup" testname="Thread Group" enabled="true">
        <stringProp name="ThreadGroup.on_sample_error">continue</stringProp>
        <elementProp name="ThreadGroup.main_controller" elementType="LoopController" guiclass="LoopControlPanel" testclass="LoopController" testname="Loop Controller" enabled="true">
          <boolProp name="LoopController.continue_forever">false</boolProp>
          <stringProp name="LoopController.loops">$Loops</stringProp>
        </elementProp>
        <stringProp name="ThreadGroup.num_threads">$threadCountRef</stringProp>
        <stringProp name="ThreadGroup.ramp_time">$rampUpRef</stringProp>
        <boolProp name="ThreadGroup.scheduler">false</boolProp>
        <stringProp name="ThreadGroup.duration"></stringProp>
        <stringProp name="ThreadGroup.delay"></stringProp>
        <boolProp name="ThreadGroup.same_user_on_next_iteration">true</boolProp>
      </ThreadGroup>
      <hashTree>
        <ConfigTestElement guiclass="HttpDefaultsGui" testclass="ConfigTestElement" testname="HTTP Request Defaults" enabled="true">
          <elementProp name="HTTPsampler.Arguments" elementType="Arguments" guiclass="HTTPArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
            <collectionProp name="Arguments.arguments"/>
          </elementProp>
          <stringProp name="HTTPSampler.domain">$baseUrlRef</stringProp>
          <stringProp name="HTTPSampler.port">443</stringProp>
          <stringProp name="HTTPSampler.protocol">https</stringProp>
          <stringProp name="HTTPSampler.contentEncoding">UTF-8</stringProp>
          <stringProp name="HTTPSampler.path"></stringProp>
          <stringProp name="HTTPSampler.connect_timeout">10000</stringProp>
          <stringProp name="HTTPSampler.response_timeout">30000</stringProp>
        </ConfigTestElement>
        <hashTree/>
        <HeaderManager guiclass="HeaderPanel" testclass="HeaderManager" testname="Auth Headers" enabled="true">
          <collectionProp name="HeaderManager.headers">
            <elementProp name="" elementType="Header">
              <stringProp name="Header.name">Authorization</stringProp>
              <stringProp name="Header.value">$authTokenRef</stringProp>
            </elementProp>
            <elementProp name="" elementType="Header">
              <stringProp name="Header.name">Accept</stringProp>
              <stringProp name="Header.value">application/json</stringProp>
            </elementProp>
            <elementProp name="" elementType="Header">
              <stringProp name="Header.name">OData-MaxVersion</stringProp>
              <stringProp name="Header.value">4.0</stringProp>
            </elementProp>
            <elementProp name="" elementType="Header">
              <stringProp name="Header.name">OData-Version</stringProp>
              <stringProp name="Header.value">4.0</stringProp>
            </elementProp>
          </collectionProp>
        </HeaderManager>
        <hashTree/>
$csvDataSetXml$samplersXml
        <ResultCollector guiclass="SummaryReport" testclass="ResultCollector" testname="Summary Report" enabled="true">
          <boolProp name="ResultCollector.error_logging">false</boolProp>
          <objProp>
            <name>saveConfig</name>
            <value class="SampleSaveConfiguration">
              <time>true</time>
              <latency>true</latency>
              <timestamp>true</timestamp>
              <success>true</success>
              <label>true</label>
              <code>true</code>
              <message>true</message>
              <threadName>true</threadName>
              <dataType>true</dataType>
              <encoding>false</encoding>
              <assertions>true</assertions>
              <subresults>true</subresults>
              <responseData>false</responseData>
              <samplerData>false</samplerData>
              <xml>false</xml>
              <fieldNames>true</fieldNames>
              <responseHeaders>false</responseHeaders>
              <requestHeaders>false</requestHeaders>
              <responseDataOnError>false</responseDataOnError>
              <saveAssertionResultsFailureMessage>true</saveAssertionResultsFailureMessage>
              <assertionsResultsToSave>0</assertionsResultsToSave>
              <bytes>true</bytes>
              <sentBytes>true</sentBytes>
              <url>true</url>
              <threadCounts>true</threadCounts>
              <idleTime>true</idleTime>
              <connectTime>true</connectTime>
            </value>
          </objProp>
          <stringProp name="filename"></stringProp>
        </ResultCollector>
        <hashTree/>
      </hashTree>
    </hashTree>
  </hashTree>
</jmeterTestPlan>
"@

$jmx | Set-Content -Path $OutputJmx -Encoding UTF8
Write-Host "$OutputJmx ($samplerCount samplers)"

if ($RecordIdGuid) {
  # Default single-row companion file so the jmx runs standalone (e.g. via
  # Run-JMeter-Tests.ps1) without needing any record-id CSV generation step;
  # Run-JMeter-Tests-Increment.ps1 overrides this with -Jrecord_ids_file for a
  # real bulk list. Written next to the jmx since CSVDataSet resolves relative
  # filenames against the running test plan's directory.
  # Split-Path -Parent returns '' (not '.') for a bare relative filename, so
  # resolve to a full path first rather than feeding '' to Join-Path.
  # GetUnresolvedProviderPathFromPSPath (not [IO.Path]::GetFullPath's 2-arg
  # overload, which is .NET Core-only and missing on Windows PowerShell 5.1)
  # resolves relative paths against the current location on both PS versions.
  $outputJmxFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJmx)
  $defaultCsv = Join-Path (Split-Path $outputJmxFullPath -Parent) 'record_ids.csv'
  if (-not (Test-Path $defaultCsv)) {
    Set-Content -Path $defaultCsv -Value @('record_id', $RecordIdGuid) -Encoding UTF8
    Write-Host "$defaultCsv (default record_id CSV, 1 row)"
  }
}
