<#
.SCRIPT_NAME User Audit Report
.SCRIPT_DESCRIPTION Pulls sign-in and license details for the selected user from Entra ID.
.REQUIRES_USER true
#>
param(
    [string]$UserPrincipalName
)

# Static app configuration is available via environment variables --
# these are the same App Service Application Settings the Node backend uses,
# inherited automatically by this process. Do NOT hardcode secrets here.
$TenantId     = $env:AzureTenantID
$ClientId     = $env:AzureAppID
$ClientSecret = $env:MICROSOFT_PROVIDER_AUTHENTICATION_SECRET

if (-not $UserPrincipalName) {
    Write-Error "No target user selected. Select a user on the dashboard before running this script."
    exit 1
}

Write-Output "Requesting Graph token..."
$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
}
$tokenResponse = Invoke-RestMethod -Method Post `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Body $tokenBody

$headers = @{ Authorization = "Bearer $($tokenResponse.access_token)" }

Write-Output "Fetching account details for $UserPrincipalName..."
$user = Invoke-RestMethod -Method Get `
    -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName`?`$select=displayName,accountEnabled,userType,createdDateTime,signInActivity" `
    -Headers $headers

Write-Output "Display Name : $($user.displayName)"
Write-Output "Enabled      : $($user.accountEnabled)"
Write-Output "Account Type : $($user.userType)"
Write-Output "Created      : $($user.createdDateTime)"
Write-Output "Audit complete."
