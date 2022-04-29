using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Interact with query parameters or the body of the request.
<# $name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name
} #>

$secretSecureString = Get-AzKeyVaultSecret -VaultName "keyvaultspauditlogs" -Name "mycert"
$secretPlainText = ConvertFrom-SecureString -AsPlainText -SecureString $secretSecureString.SecretValue
$secretPlainText

Import-Module -Name "ExchangeOnlineManagement" -RequiredVersion "2.0.5"
Connect-ExchangeOnline -CertificateThumbPrint "D7DA28894D47C9928E4E9A36CBED84015FEA4DAD" -AppID "62d46db3-d611-4781-93f6-5a1b8f0c5360" -Organization "M365x798487.onmicrosoft.com"
Import-Module -Name "PnP.PowerShell" -RequiredVersion "1.10.0"
Connect-PnPOnline -ClientId "62d46db3-d611-4781-93f6-5a1b8f0c5360" -Thumbprint "D7DA28894D47C9928E4E9A36CBED84015FEA4DAD" -Tenant "M365x798487.onmicrosoft.com" -Url "https://M365x798487-admin.sharepoint.com"

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = "Ok"
})