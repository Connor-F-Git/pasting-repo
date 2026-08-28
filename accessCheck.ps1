# --- 1. CONFIGURATION ---
$tenantName = "yourtenant" # Just the name (e.g., 'contoso' for contoso.sharepoint.com)
$siteUrl    = "https://$tenantName.sharepoint.com/sites/YourSiteName"
$listName   = "Your List Name"

# --- 2. AUTHENTICATION (Device Code Flow) ---
# Using the standard Microsoft Azure CLI App ID to request a token
$clientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
$scope    = "https://$tenantName.sharepoint.com/.default"

$deviceCodeRequest = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode" -Body @{
    client_id = $clientId
    scope     = $scope
}

Write-Host ""
Write-Host "ACTION REQUIRED: $($deviceCodeRequest.message)" -ForegroundColor Cyan
Write-Host "Waiting for you to log in... (Checking every 5 seconds)"

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
        Write-Host "Login successful! Token acquired." -ForegroundColor Green
    } catch {
        # Silent catch: We expect "authorization pending" errors until you finish logging in
    }
}

# --- 3. CREATE SHAREPOINT ITEM ---
$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/json;odata=nometadata"
    "Content-Type"  = "application/json"
}

# A basic payload testing the default 'Title' column
$body = @{
    "Title" = "Native REST API Test Item"
} | ConvertTo-Json

$postUrl = "$siteUrl/_api/web/lists/getbytitle('$listName')/items"

Write-Host "Attempting to create item in '$listName'..." -ForegroundColor Yellow

try {
    $response = Invoke-RestMethod -Uri $postUrl -Method Post -Headers $headers -Body $body
    Write-Host "Success! Test item created with ID: $($response.Id)" -ForegroundColor Green
} catch {
    Write-Host "Failed to create item. Check your permissions or list name." -ForegroundColor Red
    Write-Host $_.Exception.Message
}
