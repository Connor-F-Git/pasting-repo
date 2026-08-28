# 1. Ensure PnP 1.12.0 is loaded
Import-Module -Name PnP.PowerShell -RequiredVersion 1.12.0

$siteUrl  = "https://frbprod1.sharepoint.com/sites/SYS-FDR/STAGE"
$listName = "TriggerCopyDataverseItemToSharePoint"

# 2. Connect using browser cookie auth (bypasses App Registration)
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# 3. Add the item
$itemValues = @{
    "Title" = "Test Item via WebLogin"
}

try {
    $newItem = Add-PnPListItem -List $listName -Values $itemValues
    Write-Host "Success! Created item ID: $($newItem.Id)" -ForegroundColor Green
} catch {
    Write-Host "Failed to add item: $_" -ForegroundColor Red
}
