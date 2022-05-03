using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Interact with query parameters or the body of the request.
<# $name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name
} #>

#region Parameters
# Environment Variables
$keyVaultName = "keyvaultspauditlogs"
$keyVaultCertName = "mycert"
$appId = "62d46db3-d611-4781-93f6-5a1b8f0c5360"
$tenant = "M365x798487.onmicrosoft.com"

# Request Body Parameters
<# REQUEST BODY
    $start         # dd/MM/yyyy
    $end           # dd/MM/yyyy
    $operations    # Operation1,Operation2,OperationN
    $siteUrl       # https:\\xxx
    docLibUrl      # https:\\yyy 
#>
$startDay = "02/05/2022"
$endDay = "02/05/2022"
$operations = @("PageViewed", "FileAccessed")
$siteUrl = "https://m365x798487.sharepoint.com"
$docLibName = "Docs"
#endregion Parameters

$errorMsg = $null
$scriptStackTrace = $null
$statusCode = [HttpStatusCode]::OK
$totalCount = 0

try {
    #region Connections
    $keyVaultCertificateSecret = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $keyVaultCertName -AsPlainText
    $certificate = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ([Convert]::FromBase64String( $keyVaultCertificateSecret )), '', 'Exportable,MachineKeySet,PersistKeySet'

    Import-Module -Name "PnP.PowerShell" -RequiredVersion "1.10.0"
    Connect-PnPOnline -ClientId $appId -Tenant $tenant -Url $siteUrl -CertificateBase64Encoded $keyVaultCertificateSecret -ErrorAction Stop

    Import-Module -Name "ExchangeOnlineManagement" -RequiredVersion "2.0.5"
    Connect-ExchangeOnline -AppID $appId -Organization $tenant -Certificate $certificate
    #endregion Connections

    #region Script
    [DateTime]$start = Get-Date -Year $startDay.Split("/")[2] -Month $startDay.Split("/")[1] -Day $startDay.Split("/")[0] -Hour 0 -Minute 0 -Second 0
    [DateTime]$end = Get-Date -Year $endDay.Split("/")[2] -Month $endDay.Split("/")[1] -Day $endDay.Split("/")[0] -Hour 0 -Minute 0 -Second 0
    $end = $end.AddHours(23).AddMinutes(59).AddSeconds(59)
    $nowTime = [DateTime]::UtcNow.ToString("ddMMyyyy hhmmss")
    $outputFile = "$($(Get-Location).Path)\AuditLogRecords_$nowTime.csv"
    $record = "SharePoint"
    $resultSize = 5000
    $intervalMinutes = 720
    [DateTime]$currentStart = $start
    [DateTime]$currentEnd = $start

    $site = get-pnpsite -Includes Id -ErrorAction Stop

    Write-Host "Retrieving audit records for the date range between $($start) and $($end), RecordType=$record, ResultsSize=$resultSize"

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
    
        do {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -Operations $operations -SiteIds $site.id
        
            if (($results | Measure-Object).Count -ne 0) {
                      
                foreach ($record in $results) {
                    $json = ConvertFrom-Json $record.AuditData                
                    Add-Member -InputObject $record -MemberType NoteProperty -Name "ObjectId" -Value $json.ObjectId
                }

                $results | export-csv -Path $outputFile -Append -NoTypeInformation

                $currentTotal = $results[0].ResultCount
                $totalCount += $results.Count
            
                if ($currentTotal -eq $results[$results.Count - 1].ResultIndex) {
                    Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." -foregroundColor Yellow
                    break
                }
            }
        }
        while (($results | Measure-Object).Count -ne 0)

        $currentStart = $currentEnd
    }
    Write-Host "Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount" -foregroundColor Green

    if (Test-Path -Path $outputFile) {
        Add-PnPFile -Path $outputFile -Folder $docLibName -ErrorAction Stop
        Remove-Item -Path $outputFile
    }
    #endregion Script

    #region Disconnections
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-PnPOnline
    #endregion Disconnections
}
catch {
    $errorMsg = $_.Exception.Message
    $scriptStackTrace = $_.ScriptStackTrace
    $statusCode = [HttpStatusCode]::InternalServerError
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = $statusCode
        Body       = "{""RecordsRetrieved"" = ""$totalCount"", ""ErrorMsg"" = ""$errorMsg"",""ScriptStackTrace"" = ""$scriptStackTrace""}"
    })