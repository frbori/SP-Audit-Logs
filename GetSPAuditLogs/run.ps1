using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

#ReturnError -StatusCode "BadRequest" -ErrorDesc "ErrorDesc" -ExceptionMsg "ExceptionMsg" -ScriptStackTrace "ScriptStackTrace"

#region Checking Request Body Parameters
if ($null -eq $Request.Body.Start) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{Error = "Missing parameter (Start)." } | ConvertTo-Json
        })
    exit
}
if ($null -eq $Request.Body.End) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{Error = "Missing parameter (End)." } | ConvertTo-Json
        })
    exit
}
if ($null -eq $Request.Body.SiteUrl) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{Error = "Missing parameter (SiteUrl)." } | ConvertTo-Json
        })
    exit
}
if ($null -eq $Request.Body.User) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{Error = "Missing parameter (User)." } | ConvertTo-Json
        })
    exit
}
$startDay = $Request.Body.Start
$endDay = $Request.Body.End
$operations = $Request.Body.Operations
$siteUrl = $Request.Body.SiteUrl
$user = $Request.Body.User
#endregion Checking Request Body Parameters

#region Connections
try {
    $keyVaultCertificateSecret = Get-AzKeyVaultSecret -VaultName $env:KEY_VAULT_NAME -Name $env:KEY_VAULT_CERT_NAME -AsPlainText
}
catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error       = "An error occurred while retrieving the Azure Key Vault Secret."
                ExceptionMsg     = "$($_.Exception.Message)"
                ScriptStackTrace = "$($_.ScriptStackTrace)"
            } | ConvertTo-Json
        })
    exit
}
try {
    $certificate = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ([Convert]::FromBase64String( $keyVaultCertificateSecret )), '', 'Exportable,MachineKeySet,PersistKeySet'
}
catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error       = "An error occurred while creating the certificate from the Azure Key Vault Secret."
                ExceptionMsg     = "$($_.Exception.Message)"
                ScriptStackTrace = "$($_.ScriptStackTrace)"
            } | ConvertTo-Json
        })
    exit
}

Import-Module -Name "PnP.PowerShell" -RequiredVersion "1.10.0"
try {
    Connect-PnPOnline -ClientId $env:APP_ID -Tenant $env:TENANT -Url $siteUrl -CertificateBase64Encoded $keyVaultCertificateSecret -ErrorAction Stop
    $site = get-pnpsite -Includes Id -ErrorAction Stop
}
catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error = "An error occurred while connecting to PnP Online."
                ExceptionMsg     = "$($_.Exception.Message)"
                ScriptStackTrace = "$($_.ScriptStackTrace)"
            } | ConvertTo-Json
        })
    exit
}

$user = Get-PnPUser "i:0#.f|membership|$user" -Includes IsSiteAdmin
if ($null -eq $user) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error = "User not found." } | ConvertTo-Json
        })
    exit
}
elseif ($false -eq $user.IsSiteAdmin) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error = "User is not Site Collection Administrator." } | ConvertTo-Json
        })
    exit
}

Import-Module -Name "ExchangeOnlineManagement" -RequiredVersion "2.0.5"
try {
    Connect-ExchangeOnline -AppID $env:APP_ID -Organization $env:TENANT -Certificate $certificate
}
catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error       = "An error occurred while connecting to Exchange Online."
                ExceptionMsg     = "$($_.Exception.Message)"
                ScriptStackTrace = "$($_.ScriptStackTrace)"
            } | ConvertTo-Json
        })
    exit
}
#endregion Connections

#region Initializations
try {
    [DateTime]$start = Get-Date -Year $startDay.Split("/")[2] -Month $startDay.Split("/")[1] -Day $startDay.Split("/")[0] -Hour 0 -Minute 0 -Second 0
    [DateTime]$end = Get-Date -Year $endDay.Split("/")[2] -Month $endDay.Split("/")[1] -Day $endDay.Split("/")[0] -Hour 0 -Minute 0 -Second 0
    $end = $end.AddHours(23).AddMinutes(59).AddSeconds(59)
} 
catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error       = "An error occurred while parsing the input dates."
                ExceptionMsg     = "$($_.Exception.Message)"
                ScriptStackTrace = "$($_.ScriptStackTrace)"
            } | ConvertTo-Json
        })
    exit
}

if ($start -gt $end)
{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body       = @{Error       = "Start date can't be greater than end date."} | ConvertTo-Json
    })
    exit
}
elseif ($start -lt [DateTime]::UtcNow.AddHours(-24*90))
{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body       = @{Error       = "Start date can't be more than 90 days in the past."} | ConvertTo-Json
    })
    exit
}
elseif ((new-timespan -Start $start -End $end).TotalDays -gt 30) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body       = @{Error       = "Time interval (end - start) can't be greater than 30 days."} | ConvertTo-Json
    })
    exit
}

$nowTime = [DateTime]::UtcNow.ToString("ddMMyyyy hhmmss")
$outputFile = "$($(Get-Location).Path)\AuditLogRecords_$nowTime.csv"
$record = "SharePoint"
$resultSize = 5000
$totalCount = 0
[DateTime]$currentStart = $start
[DateTime]$currentEnd = $start
$docLibName = "SPAuditLogs_$($site.Id.ToString().split("-")[0])"
#endregion Initializations

#region Retrieving Audit Log Data
Write-Host "Retrieving audit records for the date range between $($start) and $($end), RecordType=$record, ResultsSize=$resultSize"
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
        if ($null -eq $Request.Body.Operations) {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -SiteIds $site.id
        }
        else {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -Operations $operations -SiteIds $site.id
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
                Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." -foregroundColor Yellow
                break
            }
        }
    }
    while (($results | Measure-Object).Count -ne 0)

    $currentStart = $currentEnd
}
Write-Host "Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount" -foregroundColor Green

$docLib = Get-PnPList -Identity $docLibName
if ($null -eq $docLib) {
    try {
        $docLib = New-PnPList -Title $docLibName -Template DocumentLibrary
        Set-PnPList -BreakRoleInheritance -Identity $docLib
    }
    catch {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = @{Error       = "An error occurred while creating the Document Library."
                ExceptionMsg     = "$($_.Exception.Message)"
                ScriptStackTrace = "$($_.ScriptStackTrace)"
            } | ConvertTo-Json
        })
        exit
    }
}

if (Test-Path -Path $outputFile) {
    try {
        Add-PnPFile -Path $outputFile -Folder $docLibName -ErrorAction Stop
    }
    catch {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::InternalServerError
                Body       = @{Error       = "An error occurred while uploading the report .csv file into SharePoint."
                    ExceptionMsg     = "$($_.Exception.Message)"
                    ScriptStackTrace = "$($_.ScriptStackTrace)"
                } | ConvertTo-Json
            })
        exit
    }
    Remove-Item -Path $outputFile
}
#endregion Retrieving Audit Log Data

#region Disconnections
#Disconnect-ExchangeOnline -Confirm:$false
#Disconnect-PnPOnline
Disconnect
#endregion Disconnections

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = @{RecordsRetrieved = $totalCount } | ConvertTo-Json
    })