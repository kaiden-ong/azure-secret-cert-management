# Variables are found under: automation account/shared resources/variables
$AppID = Get-AutomationVariable -Name 'appID' 
$TenantID = Get-AutomationVariable -Name 'tenantID'
$AppSecret = Get-AutomationVariable -Name 'appSecret'  


[int32] $expirationDays = 90 # Finds secrets/certs expiring within this many days
[string] $emailSender = "kaiden.ong000@gmail.com"
[string[]]$emailTo = ,"kaiden.ong000@gmail.com" # To add more just separate with commas


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
        [system.string]$Subject = "App Secret Expiration Notice",
        [system.string]$Body
    )
    begin {
        $headers = @{
            "Authorization" = "Bearer $($AccessToken)"
            "Content-type"  = "application/json"
        }


        $Recipients = $To | ForEach-Object {
            @{
                "emailAddress" = @{
                    "address" = $_
                }
            }
        }


        $BodyJsonsend = @"
{
   "message": {
        "subject": "$Subject",
        "body": {
            "contentType": "HTML",
            "content": "$($Body)"
        },
        "toRecipients": $($Recipients | ConvertTo-Json -Compress)
   },
   "saveToSentItems": "true"
}
"@
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
$apps = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/" 
$allApps = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/servicePrincipals"

# Loops through each app and looks for secrets and certificates
foreach ($app in $apps) {
    # Within each app look at each secret
    $app.passwordCredentials | foreach-object {
        # Adds to array if the secret has an expiration date within $expirationDays (90) days
        if ($_.endDateTime -ne $null) {
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

# Filter for enterprise apps only
$enterpriseApps = @()
foreach ($app in $allApps) {
    if ($app.Tags -contains "WindowsAzureActiveDirectoryIntegratedApp") {
        $enterpriseApps += $app
    }
} 

foreach ($app in $enterpriseApps) {
    $currSsoCertArray = @()
    $hasValidCert = $false
    $signVerifyDup = @{}
    $app.keyCredentials | foreach-object {
        # Adds to array if the secret has an expiration date within $expirationDays (90) days
        if ($_.endDateTime -ne $null) {
            [system.string]$ssoCertificateDisplayName = $_.displayName
            [system.string]$id = $app.id
            [system.string]$displayName = $app.appDisplayName
            $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.endDateTime, 'Pacific Standard Time')
            [int32]$daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Pacific Standard Time")) -End $Date).Days
            
            if (($daysUntilExpiration -ne $null) -and ($daysUntilExpiration -le $expirationDays)) {
                if (-not $signVerifyDup.Contains($_.customKeyIdentifier)) {
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
                }
            } else {
                $hasValidCert = $true
            }
            $daysUntilExpiration = $null
            $ssoCertificateDisplayName = $null
        }
    }
    if (-not $hasValidCert) {
        $ssoCertificateArray += $currSsoCertArray
    }
}

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
<h1>Weekly Secret & Certificate Report</h1>
<p>Below are the secrets and certificates set to expire in 90 days. If an application has a secret that is within 90 days
of expiring, but also one that expires later, then it will not be listed.<p>
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
Send-MSGraphEmail -Uri "https://graph.microsoft.com/v1.0/users/$emailSender/sendMail" -AccessToken $tokenResponse.access_token -To $emailTo -Body $combinedTable
