# 1. Clean out any modern versions that won't load on PS 5.1
Uninstall-Module -Name PnP.PowerShell -AllVersions -Force -ErrorAction SilentlyContinue

# 2. Install the PS 5.1 compatible version (v1.12.0)
Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser -AllowClobber -Force

# 3. Explicitly import v1.12.0 into your session
Import-Module -Name PnP.PowerShell -RequiredVersion 1.12.0

# 4. Connect to your site
$siteUrl = "https://YOUR_TENANT.sharepoint.com/sites/SYS-FDR/STAGE"
Connect-PnPOnline -Url $siteUrl -Interactive
