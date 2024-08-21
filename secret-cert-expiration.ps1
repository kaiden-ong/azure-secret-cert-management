# 2024-08-20 pro edition
# Variables are found under: automation accounts/mhs-shd-uw2-infr-appsecretexp-aa-001/shared resources/variables
$AppID = Get-AutomationVariable -Name 'appID' 
$TenantID = Get-AutomationVariable -Name 'tenantID'
$AppSecret = Get-AutomationVariable -Name 'appSecret'  


[int32] $expirationDays = 90 # Finds secrets/certs expiring within this many days
[string] $emailSender = "kaiden.ong000@gmail.com"
[string[]]$emailTo = "kaiden.ong000@gmail.com" # To add more just separate with commas


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
$apps = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/" 


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
                $certificateArray += $_ | Select-Object @{
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
            }
            $daysUntilExpiration = $null
            $certificateDisplayName = $null
        }
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
        padding: 5px 10px;
        border-bottom: 1px solid #d1d1d1;
    }
</style>
"@


# Generate HTML tables without the -Head parameter
if ($secretArray -ne 0) {
    $secretTable = $secretArray | Sort-Object "Days Until Expiration" | Select-Object "Application Name", "Secret Name", "Days Until Expiration" | ConvertTo-Html -Fragment
} else {
    $secretTable = "<p>No apps with expiring secrets</p>"
}


if ($certificateArray -ne 0) {
    $certificateTable = $certificateArray | Sort-Object "Days Until Expiration" | Select-Object "Application Name", "Certificate Name", "Days Until Expiration" | ConvertTo-Html -Fragment
} else {
    $certificateTable = "<p>No apps with expiring certificates</p>"
}


# Combine tables with styling
$combinedTable = @"
<html>
<head>
$style
</head>
<body>
<h2>Expiring Secrets</h2>$secretTable<h2>Expiring Certificates</h2>$certificateTable
</body>
</html>
"@


write-output "sending email"


write-output $emailTo
Send-MSGraphEmail -Uri "https://graph.microsoft.com/v1.0/users/$emailSender/sendMail" -AccessToken $tokenResponse.access_token -To $emailTo -Body $combinedTable