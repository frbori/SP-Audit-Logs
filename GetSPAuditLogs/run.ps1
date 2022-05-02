using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Interact with query parameters or the body of the request.
<# $name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name
} #>

#region Parameters
$keyVaultName = "keyvaultspauditlogs"
$keyVaultCertName = "mycert"
$appId = "62d46db3-d611-4781-93f6-5a1b8f0c5360"
$tenant = "M365x798487.onmicrosoft.com"
$siteUrl = "https://M365x798487.sharepoint.com"
#endregion Parameters

#region Connections
$keyVaultCertificateSecret = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $keyVaultCertName -AsPlainText
$certificate = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ([Convert]::FromBase64String( $keyVaultCertificateSecret )), '', 'Exportable,MachineKeySet,PersistKeySet'

Import-Module -Name "PnP.PowerShell" -RequiredVersion "1.10.0"
Connect-PnPOnline -ClientId $appId -Tenant $tenant -Url $siteUrl -CertificateBase64Encoded $keyVaultCertificateSecret

Import-Module -Name "ExchangeOnlineManagement" -RequiredVersion "2.0.5"
Connect-ExchangeOnline -AppID $appId -Organization $tenant -Certificate $certificate
#endregion Connections

#Get-Location    # C:\home\site\wwwroot

#region Script temp
[DateTime]$start = [DateTime]::UtcNow.AddHours(-24)
[DateTime]$end = [DateTime]::UtcNow

$nowTime = [DateTime]::UtcNow.ToString("ddMMyyyy hhmmss")
$outputFile = "$($(Get-Location).Path)\AuditLogRecords_$nowTime.csv"
$record = "SharePoint"
$resultSize = 5000
$intervalMinutes = 60
[DateTime]$currentStart = $start
[DateTime]$currentEnd = $start

Write-Host "Retrieving audit records for the date range between $($start) and $($end), RecordType=$record, ResultsSize=$resultSize"

$totalCount = 0
while ($true) {
    $currentEnd = $currentStart.AddMinutes($intervalMinutes)
    if ($currentEnd -gt $end) {
        $currentEnd = $end
    }

    if ($currentStart -eq $currentEnd) {
        break
    }

    $sessionID = [Guid]::NewGuid().ToString() + "_" + "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
    Write-Host "Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    $currentCount = 0

    do {
        $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -ObjectIds "$siteUrl/*" -Operations PageViewed
        
        if (($results | Measure-Object).Count -ne 0) {
            $results | export-csv -Path $outputFile -Append -NoTypeInformation    
                        
            foreach ($record in $results) {
                $json = ConvertFrom-Json $record.AuditData                
                
                Write-Host "User:" $record.UserIds     -ForegroundColor DarkCyan
                Write-Host "On:" $record.CreationDate  -ForegroundColor DarkCyan
                Write-Host "Visited:" $json.ObjectId   -ForegroundColor DarkCyan 
                Write-host "----"
            }

            $currentTotal = $results[0].ResultCount
            $totalCount += $results.Count
            $currentCount += $results.Count
            
            if ($currentTotal -eq $results[$results.Count - 1].ResultIndex) {
                Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." -foregroundColor Yellow
                break
            }
        }
    }
    while (($results | Measure-Object).Count -ne 0)

    $currentStart = $currentEnd
}
Write-Host "Script complete! Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount" -foregroundColor Green
#endregion Script temp

#region Disconnections
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-PnPOnline
#endregion Disconnections

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = "Ok"
})