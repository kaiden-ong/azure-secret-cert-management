# Variables are found under: automation account/shared resources/variables
$AppID = Get-AutomationVariable -Name 'appID' 
$TenantID = Get-AutomationVariable -Name 'tenantID'
$AppSecret = Get-AutomationVariable -Name 'appSecret'  


[int32] $expirationDays = 90 # Finds secrets/certs expiring within this many days
[string] $emailSender = "kaiden.ong000@gmail.com"
[string[]]$emailTo = @("kaiden.ong@gmail.com") # To add more just separate with commas


# Establishes connection to the MS Graph API, returning a token that gives access to Azure resources 
Function Connect-MSGraphAPI {
    param (
        [system.string]$AppID,
        [system.string]$TenantID,
        [system.string]$AppSecret
    )
    begin {
        $URI = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        $ReqTokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $AppID
            Client_Secret = $AppSecret
        } 
    }
    Process {
        Write-Host "Connecting to the Graph API"
        $Response = Invoke-RestMethod -Uri $URI -Method POST -Body $ReqTokenBody
    }
    End {
        $Response
    }
}


# Recursively gets all app registrations
Function Get-MSGraphRequest {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken
    )
    begin {
        [System.Array]$allPages = @()
        $ReqTokenBody = @{
            Headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $($AccessToken)"
            }
            Method  = "Get"
            Uri     = $Uri
        }
    }
    process {
        write-verbose "GET request at endpoint: $Uri"
        $data = Invoke-RestMethod @ReqTokenBody
        while ($data.'@odata.nextLink') {
            $allPages += $data.value
            $ReqTokenBody.Uri = $data.'@odata.nextLink'
            $Data = Invoke-RestMethod @ReqTokenBody
            # to avoid throttling, the loop will sleep for 3 seconds
            Start-Sleep -Seconds 3
        }
        $allPages += $data.value
    }
    end {
        Write-Verbose "Returning all results"
        $allPages
    }
}


# Creates and sends email
Function Send-MSGraphEmail {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken,
        [system.string[]]$To,
        [system.string]$Subject = "App Secret/Certificate Expiration Notice",
        [system.string]$Body,
        [System.Collections.Hashtable[]]$Attachment
    )
    begin {
        $headers = @{
            "Authorization" = "Bearer $($AccessToken)"
            "Content-type"  = "application/json"
        }
        
        if ($To.length -eq 1) {
            $Recipients = @(
                @{
                    "emailAddress" = @{
                        "address" = $To[0]
                    }
                }
            )
        } else {
            $Recipients = $To | ForEach-Object {
                @{
                    "emailAddress" = @{
                        "address" = $_
                    }
                }
            }
        }

        $BodyContent = @{
            "message" = @{
                "subject" = $Subject
                "body" = @{
                    "contentType" = "HTML"
                    "content" = $Body
                }
                "toRecipients" = $Recipients
                "attachments" = @($Attachment)
            }
            "saveToSentItems" = "true"
        }

        $BodyJsonsend = $BodyContent | ConvertTo-Json -Depth 5
    }
    process {
        $data = Invoke-RestMethod -Method POST -Uri $Uri -Headers $headers -Body $BodyJsonsend
    }
    end {
        $data
    }
}

$tokenResponse = Connect-MSGraphAPI -AppID $AppID -TenantID $TenantID -AppSecret $AppSecret

$secretArray = @()
$certificateArray = @()
$ssoCertificateArray = @()

$secretData = @()
$certData = @()
$ssoData = @()

$apps = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/" 
$allApps = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/servicePrincipals"

# Loops through each app and looks for secrets and certificates
foreach ($app in $apps) {
    # Within each app look at each secret
    $app.passwordCredentials | foreach-object {
        # Adds to array if the secret has an expiration date within $expirationDays (90) days
        if ($_.endDateTime -ne $null) {
            # For email secret.csv file
            $secretData += @{
                Type                  = "App Registration Secret"
                ID                    = $app.id ?? ""
                App_Name              = $app.displayName ?? ""
                Secret_ID             = $_.keyId ?? ""
                Secret_Name           = $_.displayName ?? ""
                Creation_Date         = $_.startDateTime ?? ""
                Expiration_Date       = $_.endDateTime ?? ""
            }

            # For email html table
            [system.string]$secretDisplayName = $_.displayName
            [system.string]$id = $app.id
            [system.string]$displayName = $app.displayName
            $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.endDateTime, 'Pacific Standard Time')
            [int32]$daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Pacific Standard Time")) -End $Date).Days
            
            if (($daysUntilExpiration -ne $null) -and ($daysUntilExpiration -le $expirationDays)) {
                $secretArray += $_ | Select-Object @{
                    name = "id"; 
                    expr = { $id } 
                }, 
                @{
                    name = "Application Name"; 
                    expr = { $displayName } 
                }, 
                @{
                    name = "Secret Name"; 
                    expr = { $secretDisplayName } 
                },
                @{
                    name = "Days Until Expiration"; 
                    expr = { $daysUntilExpiration } 
                }
            }
            $daysUntilExpiration = $null
            $secretDisplayName = $null
        }
    }

    $currCertArray = @()
    $hasValidCert = $false
    # Within each app look at each certificate
    $app.keyCredentials | foreach-object {
        # Adds to array if the certificate has an expiration date within $expirationDays (90) days
        if ($_.endDateTime -ne $null) {
            # For email certData.csv file
            $certData += @{
                Type                  = "App Registration Certificate"
                ID                    = $app.id ?? ""
                App_Name              = $app.displayName ?? ""
                Cert_ID               = $_.keyId ?? ""
                Cert_Name             = $_.displayName ?? ""
                Thumbprint            = $app.customKeyIdentifier ?? ""
                Creation_Date         = $_.startDateTime ?? ""
                Expiration_Date       = $_.endDateTime ?? ""
            }

            # For email html table
            [system.string]$certificateDisplayName = $_.displayName
            [system.string]$id = $app.id    
            [system.string]$displayName = $app.displayName
            $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.endDateTime, 'Pacific Standard Time')
            [int32]$daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Pacific Standard Time")) -End $Date).Days
            
            if (($daysUntilExpiration -ne $null) -and ($daysUntilExpiration -le $expirationDays)) {
                $currCertArray += $_ | Select-Object @{
                    name = "id"; 
                    expr = { $id } 
                }, 
                @{
                    name = "Application Name"; 
                    expr = { $displayName } 
                }, 
                @{
                    name = "Certificate Name"; 
                    expr = { $certificateDisplayName } 
                },
                @{
                    name = "Days Until Expiration"; 
                    expr = { $daysUntilExpiration } 
                }
            } else {
                $hasValidCert = $true
            }
            $daysUntilExpiration = $null
            $certificateDisplayName = $null
        }
    }
    if (-not $hasValidCert) {
        $certificateArray += $currCertArray
    }
}




# foreach ($app in $apps) {
#     $app.passwordCredentials | foreach-object {
#         if ($_.endDateTime -ne $null) {
#             $secretData += @{
#                 Type                  = "App Registration Secret"
#                 ID                    = $app.id ?? ""
#                 App_Name              = $app.displayName ?? ""
#                 Secret_ID             = $_.keyId ?? ""
#                 Secret_Name           = $_.displayName ?? ""
#                 Creation_Date         = $_.startDateTime ?? ""
#                 Expiration_Date       = $_.endDateTime ?? ""
#             }
#         }
#     }

#     $app.keyCredentials | foreach-object {
#         if ($_.endDateTime -ne $null) {
#             $certData += @{
#                 Type                  = "App Registration Certificate"
#                 ID                    = $app.id ?? ""
#                 App_Name              = $app.displayName ?? ""
#                 Cert_ID               = $_.keyId ?? ""
#                 Cert_Name             = $_.displayName ?? ""
#                 Thumbprint            = $app.customKeyIdentifier ?? ""
#                 Creation_Date         = $_.startDateTime ?? ""
#                 Expiration_Date       = $_.endDateTime ?? ""
#             }
#         }
#     }
# }

$enterpriseApps = @()

foreach ($app in $allApps) {
    if ($app.Tags -contains "WindowsAzureActiveDirectoryIntegratedApp") {
        $enterpriseApps += $app
    }
}

# foreach ($app in $enterpriseApps) {
#     $signVerifyDup = @{}
#     $app.keyCredentials | foreach-object {
#         if (-not $signVerifyDup.Contains($_.customKeyIdentifier) -and $_.endDateTime -ne $null) {
#             $ssoData += @{
#                 Type                  = "Enterprise App SSO Certificate"
#                 ID                    = $app.id ?? ""
#                 App_Name              = $app.appDisplayName ?? ""
#                 Account_Enabled       = $app.accountEnabled ?? ""
#                 SSO_Mode              = $app.preferredSingleSignOnMode ?? ""
#                 Thumbprint            = $app.preferredTokenSigningKeyThumbprint ?? ""
#                 Notification_Emails   = $app.notificationEmailAddresses ?? ""
#                 Cert_ID               = $_.customKeyIdentifier ?? ""
#                 Cert_Name             = $_.displayName ?? ""
#                 Creation_Date         = $_.startDateTime ?? ""
#                 Expiration_Date       = $_.endDateTime ?? ""
#             }
#             $signVerifyDup[$_.customKeyIdentifier] = $true
#         }
#     }
# }

foreach ($app in $enterpriseApps) {
    $currSsoCertArray = @()
    $hasValidCert = $false
    $signVerifyDup = @{}
    $signVerifyDupCSV = @{}
    
    $app.keyCredentials | foreach-object {
        if (-not $signVerifyDupCSV.Contains($_.customKeyIdentifier)) {
            # For email sso.csv file
            $ssoData += @{
                Type                  = "Enterprise App SSO Certificate"
                ID                    = $app.id ?? ""
                App_Name              = $app.appDisplayName ?? ""
                Account_Enabled       = $app.accountEnabled ?? ""
                SSO_Mode              = $app.preferredSingleSignOnMode ?? ""
                Thumbprint            = $app.preferredTokenSigningKeyThumbprint ?? ""
                Notification_Emails   = $app.notificationEmailAddresses ?? ""
                Cert_ID               = $_.customKeyIdentifier ?? ""
                Cert_Name             = $_.displayName ?? ""
                Creation_Date         = $_.startDateTime ?? ""
                Expiration_Date       = $_.endDateTime ?? ""
            }
            $signVerifyDupCSV[$_.customKeyIdentifier] = $true
        }
        if ($app.accountEnabled) {
            # Adds to array if the secret has an expiration date within $expirationDays (90) days
            if ($_.endDateTime -ne $null -and -not $signVerifyDup.Contains($_.customKeyIdentifier)) {
                # For email html table
                [system.string]$ssoCertificateDisplayName = $_.displayName
                [system.string]$id = $app.id
                [system.string]$displayName = $app.appDisplayName
                $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.endDateTime, 'Pacific Standard Time')
                [int32]$daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Pacific Standard Time")) -End $Date).Days
                
                if (($daysUntilExpiration -ne $null) -and ($daysUntilExpiration -le $expirationDays)) {
                    $currSsoCertArray += $_ | Select-Object @{
                        name = "id"; 
                        expr = { $id } 
                    }, 
                    @{
                        name = "Application Name"; 
                        expr = { $displayName } 
                    }, 
                    @{
                        name = "SSO Certificate Name"; 
                        expr = { $ssoCertificateDisplayName } 
                    },
                    @{
                        name = "Days Until Expiration"; 
                        expr = { $daysUntilExpiration } 
                    }

                    # Mark this certificate as seen
                    $signVerifyDup[$_.customKeyIdentifier] = $true
                } else {
                    $hasValidCert = $true
                }
            }
        }
    }
    if (-not $hasValidCert) {
        $ssoCertificateArray += $currSsoCertArray
    }
}

# Create csv for secret array
$secretOrderedData = $secretData | Select-Object Type, ID, App_Name, Secret_ID, Secret_Name, Creation_Date, Expiration_Date
$secretCsv = $secretOrderedData | ConvertTo-Csv -NoTypeInformation
$secretCsvString = $secretCsv -join "`r`n"
$secretAttachmentContent = [System.Text.Encoding]::UTF8.GetBytes($secretCsvString)
$secretBase64Content = [System.Convert]::ToBase64String($secretAttachmentContent)

# Create csv for cert array
$certOrderedData = $certData | Select-Object Type, ID, App_Name, Cert_ID, Cert_Name, Thumbprint, Creation_Date, Expiration_Date
$certCsv = $certOrderedData | ConvertTo-Csv -NoTypeInformation
$certCsvString = $certCsv -join "`r`n"
$certAttachmentContent = [System.Text.Encoding]::UTF8.GetBytes($certCsvString)
$certBase64Content = [System.Convert]::ToBase64String($certAttachmentContent)

# Create csv for sso array
$ssoOrderedData = $ssoData | Select-Object Type, ID, App_Name, Account_Enabled, SSO_Mode, Thumbprint, Notification_Emails, Cert_ID, Cert_Name, Creation_Date, Expiration_Date
$ssoCsv = $ssoOrderedData | ConvertTo-Csv -NoTypeInformation
$ssoCsvString = $ssoCsv -join "`r`n"
$ssoAttachmentContent = [System.Text.Encoding]::UTF8.GetBytes($ssoCsvString)
$ssoBase64Content = [System.Convert]::ToBase64String($ssoAttachmentContent)

# Add all attachments as hashtables into array
$Attachments = @(
    @{
        "@odata.type" = "#microsoft.graph.fileAttachment"
        "name" = "secretData.csv"
        "contentType" = "text/csv"
        "contentBytes" = $secretBase64Content
    },
    @{
        "@odata.type" = "#microsoft.graph.fileAttachment"
        "name" = "certData.csv"
        "contentType" = "text/csv"
        "contentBytes" = $certBase64Content
    },
    @{
        "@odata.type" = "#microsoft.graph.fileAttachment"
        "name" = "ssoData.csv"
        "contentType" = "text/csv"
        "contentBytes" = $ssoBase64Content
    }
)

# Define styles for the tables
$style = @"
<style>
    table {
        border-collapse: separate;
        border-spacing: 0;
        border: 2px solid #009879;
        border-radius: 8px;
        text-align: left;
        width: 100%;
        margin: 7px;
        font-size: 0.9em;
        font-family: sans-serif;
        min-width: 400px;
        box-shadow: 0 3px 5px;
    }
    th {
        background-color: #009879;
        color: #ffffff;
        border: none;
    }
    td, th {
        padding: 3px 8px;
        border-bottom: 1px solid #d1d1d1;
    }
</style>
"@

# Generate HTML tables without the -Head parameter
if ($secretArray -ne 0) {
    $secretTable = $secretArray | Sort-Object "Days Until Expiration" | Select-Object "Application Name", "Secret Name", "Days Until Expiration" | ConvertTo-Html -Fragment
} else {
    $secretTable = "<p>No apps registrations with expiring secrets</p>"
}

if ($certificateArray -ne 0) {
    $certificateTable = $certificateArray | Sort-Object "Days Until Expiration" | Select-Object "Application Name", "Certificate Name", "Days Until Expiration" | ConvertTo-Html -Fragment
} else {
    $certificateTable = "<p>No app registrations with expiring certificates</p>"
}

if ($ssoCertificateArray -ne 0) {
    $ssoCertificateTable = $ssoCertificateArray | Sort-Object "Days Until Expiration" | Select-Object "Application Name", "SSO Certificate Name", "Days Until Expiration" | ConvertTo-Html -Fragment
} else {
    $ssoCertificateTable = "<p>No enterprise apps with expiring certificates</p>"
}

$secretCount = $secretArray.Length
$certCount = $certificateArray.Length
$ssoCertCount = $ssoCertificateArray.Length

# Combine tables with styling
$combinedTable = @"
<html>
<head>
$style
</head>
<body>
<h1>Weekly AR Secret & Certificate Report</h1>
<p>Below are the secrets and certificates set to expire in 90 days. Additionally,
the three attached csv files include all secrets, certs, and sso certs with expiration dates.<p>
<h2>Expiring Secrets</h2>
<h3>Count: $secretCount</h3>
$secretTable
<h2>Expiring Certificates</h2>
<h3>Count: $certCount</h3>
$certificateTable
<h2>Expiring Enterprise App Certificates (SSO)</h2>
<h3>Count: $ssoCertCount</h3>
$ssoCertificateTable
</body>
</html>
"@

write-output "sending email"

write-output $emailTo
# Send-MSGraphEmail -Uri "https://graph.microsoft.com/v1.0/users/$emailSender/sendMail" -AccessToken $tokenResponse.access_token -To $emailTo -Body $combinedTable -AttachmentPath
Send-MSGraphEmail -Uri "https://graph.microsoft.com/v1.0/users/$emailSender/sendMail" -AccessToken $tokenResponse.access_token -To $emailTo -Body $combinedTable -Attachment $Attachments
