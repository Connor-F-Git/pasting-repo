param(
    [Parameter(Mandatory)]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$QueryUrl
)

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
        elseif ($err -eq 'authorization_declined') { throw 'Login was declined.' }
        elseif ($err -eq 'expired_token') { throw 'Device code expired.' }
        else { throw $_ }
    }
}

Write-Host "Calling: $QueryUrl"
Write-Host ''

$response = Invoke-RestMethod `
    -Method GET `
    -Uri $QueryUrl `
    -Headers @{ Authorization = "Bearer $accessToken"; Accept = 'application/json' }

$response | ConvertTo-Json -Depth 20
