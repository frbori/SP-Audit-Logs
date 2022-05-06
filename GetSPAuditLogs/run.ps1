using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

#region Checking Request Body Parameters
if ($null -eq $Request.Body.Start) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "Missing parameter (Start)."
}
if ($null -eq $Request.Body.End) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "Missing parameter (End)."
}
if ($null -eq $Request.Body.SiteUrl) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "Missing parameter (SiteUrl)."
}
if ($null -eq $Request.Body.User) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "Missing parameter (User)."
}
#endregion Checking Request Body Parameters

#region Initializations
# Variables from request body
$startDay = $Request.Body.Start
$endDay = $Request.Body.End
$operations = $Request.Body.Operations
$siteUrl = $Request.Body.SiteUrl
$user = $Request.Body.User
# Parsing and checking Dates
try {
    [DateTime]$start = Get-Date -Year $startDay.Split("/")[2] -Month $startDay.Split("/")[1] -Day $startDay.Split("/")[0] -Hour 0 -Minute 0 -Second 0
    [DateTime]$end = Get-Date -Year $endDay.Split("/")[2] -Month $endDay.Split("/")[1] -Day $endDay.Split("/")[0] -Hour 0 -Minute 0 -Second 0
    $end = $end.AddHours(23).AddMinutes(59).AddSeconds(59)
} 
catch {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "An error occurred while parsing the input dates." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace
}
if ($start -gt $end) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "Start date can't be greater than end date."
}
elseif ($start -lt [DateTime]::UtcNow.AddHours(-24 * 90)) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "Start date can't be more than 90 days in the past."
}
elseif ((new-timespan -Start $start -End $end).TotalDays -gt 30) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "Time interval (end - start) can't be greater than 30 days."
}
# Other variables
$nowTime = [DateTime]::UtcNow.ToString("ddMMyyyy hhmmss")
$outputFile = "$($(Get-Location).Path)\AuditLogRecords_$nowTime.csv"
$record = "SharePoint"
$resultSize = 5000
$totalCount = 0
[DateTime]$currentStart = $start
[DateTime]$currentEnd = $start
#endregion Initializations

#region Connections
# Retrieving Certificate from Azure Key Vault
try {
    $keyVaultCertificateSecret = Get-AzKeyVaultSecret -VaultName $env:KEY_VAULT_NAME -Name $env:KEY_VAULT_CERT_NAME -AsPlainText
}
catch {
    ReturnError -StatusCode "InternalServerError" -ErrorDesc "An error occurred while retrieving the Azure Key Vault Secret." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace
}
try {
    $certificate = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ([Convert]::FromBase64String( $keyVaultCertificateSecret )), '', 'Exportable,MachineKeySet,PersistKeySet'
}
catch {
    ReturnError -StatusCode "InternalServerError" -ErrorDesc "An error occurred while creating the certificate from the Azure Key Vault Secret." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace
}
# Connecting to PnP
Import-Module -Name "PnP.PowerShell" -RequiredVersion "1.10.0"
try {
    Connect-PnPOnline -ClientId $env:APP_ID -Tenant $env:TENANT -Url $siteUrl -CertificateBase64Encoded $keyVaultCertificateSecret -ErrorAction Stop
    $site = get-pnpsite -Includes Id -ErrorAction Stop
    $docLibName = "SPAuditLogs_$($site.Id.ToString().split("-")[0])"
}
catch {
    ReturnError -StatusCode "InternalServerError" -ErrorDesc "An error occurred while connecting to PnP Online." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace -DisconnectPnP
}
# Checking if user exists and is Site Coll Admin
$user = Get-PnPUser "i:0#.f|membership|$user" -Includes IsSiteAdmin
if ($null -eq $user) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "User not found." -DisconnectPnP
}
elseif ($false -eq $user.IsSiteAdmin) {
    ReturnError -StatusCode "BadRequest" -ErrorDesc "User is not Site Collection Administrator." -DisconnectPnP
}
# Connecting to EXO
Import-Module -Name "ExchangeOnlineManagement" -RequiredVersion "2.0.5"
try {
    Connect-ExchangeOnline -AppID $env:APP_ID -Organization $env:TENANT -Certificate $certificate -ShowBanner:$false
}
catch {
    ReturnError -StatusCode "InternalServerError" -ErrorDesc "An error occurred while connecting to Exchange Online." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace -DisconnectPnP
}
#endregion Connections

#region Retrieving Audit Log Data
Write-Host "Retrieving audit records for the date range between $start and $end, RecordType=$record, ResultsSize=$resultSize"
while ($true) {
    $currentEnd = $currentStart.AddMinutes($env:INTERVAL_MINUTES)
    if ($currentEnd -gt $end) {
        $currentEnd = $end
    }

    if ($currentStart -eq $currentEnd) {
        break
    }

    $sessionID = [Guid]::NewGuid().ToString() + "_" + "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
    Write-Host "Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    
    do {
        try {
            if ($null -eq $Request.Body.Operations) {
                $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -SiteIds $site.id
            }
            else {
                $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -Operations $operations -SiteIds $site.id
            }
        }
        catch {
            ReturnError -StatusCode "InternalServerError" -ErrorDesc "An error occurred while retrieving the Audit Log." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace -DisconnectPnP -DisconnectEXO
        }

        if (($results | Measure-Object).Count -ne 0) {
                      
            foreach ($record in $results) {
                $json = ConvertFrom-Json $record.AuditData                
                Add-Member -InputObject $record -MemberType NoteProperty -Name "ObjectId" -Value $json.ObjectId
            }

            $results | export-csv -Path $outputFile -Append -NoTypeInformation

            $currentTotal = $results[0].ResultCount
            $totalCount += $results.Count
            
            if ($currentTotal -eq $results[$results.Count - 1].ResultIndex) {
                Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval."
                break
            }
        }
    }
    while (($results | Measure-Object).Count -ne 0)

    $currentStart = $currentEnd
}
Write-Host "Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount"
#endregion Retrieving Audit Log Data

#region Saving the report .csv file into SharePoint
$docLib = Get-PnPList -Identity $docLibName
if ($null -eq $docLib) {
    try {
        $docLib = New-PnPList -Title $docLibName -Template DocumentLibrary
        Set-PnPList -BreakRoleInheritance -Identity $docLib
    }
    catch {
        ReturnError -StatusCode "InternalServerError" -ErrorDesc "An error occurred while creating the Document Library." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace -DisconnectPnP -DisconnectEXO
    }
}
if (Test-Path -Path $outputFile) {
    try {
        Add-PnPFile -Path $outputFile -Folder $docLibName -ErrorAction Stop
    }
    catch {
        ReturnError -StatusCode "InternalServerError" -ErrorDesc "An error occurred while uploading the report .csv file into SharePoint." -ExceptionMsg $_.Exception.Message -ScriptStackTrace $_.ScriptStackTrace -DisconnectPnP -DisconnectEXO
    }
    Remove-Item -Path $outputFile
}
#endregion Saving the report .csv file into SharePoint

Disconnect

# Associate values to output bindings
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = @{RecordsRetrieved = $totalCount } | ConvertTo-Json
    })