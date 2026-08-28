# 1. Install module (run once if not already installed)
# Install-Module -Name PnP.PowerShell -Scope CurrentUser

# 2. Connect to the subsite (launches a browser prompt for authentication)
$siteUrl = "https://YOUR_TENANT.sharepoint.com/sites/SYS-FDR/STAGE"
Connect-PnPOnline -Url $siteUrl -Interactive

# 3. Target the list and set item values
$listName = "TriggerCopyDataverseItemToSharePoint"

$itemValues = @{
    "Title" = "Test Item via PnP"
}

# 4. Create the item
try {
    $newItem = Add-PnPListItem -List $listName -Values $itemValues
    Write-Host "Success! Created item with ID: $($newItem.Id)" -ForegroundColor Green
} catch {
    Write-Host "Error creating item: $_" -ForegroundColor Red
}
