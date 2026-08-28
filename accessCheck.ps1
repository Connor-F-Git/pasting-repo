# 1. CONFIGURATION
$tenantName = "YOUR_TENANT_NAME" 
$siteUrl    = "https://$tenantName.sharepoint.com/sites/SYS-FDR/STAGE"
$listUrl    = "/sites/SYS-FDR/STAGE/Lists/TriggerCopyDataverseItemToSharePoint"

# Standard PnP Management Shell Client ID used to trigger device code auth
$clientId   = "31359c7f-bd7e-475c-86b6-3e148d28200e"
$scope      = "https://$tenantName.sharepoint.com/.default"

# 2. REQUEST DEVICE CODE FOR AUTHENTICATION
$deviceCodeRequest = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode" -Body @{
    client_id = $clientId
    scope     = $scope
}

Write-Host ""
Write-Host "ACTION REQUIRED: $($deviceCodeRequest.message)" -ForegroundColor Cyan
Write-Host "Waiting for login completion..."

# 3. POLL FOR ACCESS TOKEN
$token = $null
while ($null -eq $token) {
    Start-Sleep -Seconds 5
    try {
        $tokenRequest = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/token" -Body @{
            grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
            client_id   = $clientId
            device_code = $deviceCodeRequest.device_code
        } -ErrorAction Stop
        $token = $tokenRequest.access_token
        Write-Host "Authenticated successfully!" -ForegroundColor Green
    } catch {
        # Loop continues silently until user completes browser login
    }
}

# 4. CREATE SHAREPOINT LIST ITEM
$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/json;odata=nometadata"
    "Content-Type"  = "application/json"
}

$body = @{
    "Title" = "Native REST API Test Item"
} | ConvertTo-Json

$postUrl = "$siteUrl/_api/web/GetList('$listUrl')/items"

Write-Host "Adding item to SharePoint list..." -ForegroundColor Yellow

try {
    $response = Invoke-RestMethod -Uri $postUrl -Method Post -Headers $headers -Body $body
    Write-Host "Success! Created list item with ID: $($response.Id)" -ForegroundColor Green
} catch {
    Write-Host "Failed to create item:" -ForegroundColor Red
    Write-Host $_.Exception.Message
}
