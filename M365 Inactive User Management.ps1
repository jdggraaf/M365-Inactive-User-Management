<#
.SYNOPSIS
    Disables inactive Microsoft 365 user accounts across multiple tenants and sends a report.

.DESCRIPTION
    This script connects to Microsoft Graph using an application (client) credential flow.
    It retrieves all customers (tenants) under the main tenant via Delegated Administration.
    It then identifies and disables accounts that are deemed inactive, excluding those
    in a specified "exception" group.

    Finally, it emails a log of the actions taken to specified recipients.

.NOTES
    Author:      Joost de Graaf
    Version:     1.0
    Date:        2025-01-01

.PARAMETER ApplicationID
    The Azure AD Application (Client) ID used to authenticate with Microsoft Graph.

.PARAMETER ApplicationSecret
    The Azure AD Application (Client) Secret used to authenticate with Microsoft Graph.

.PARAMETER MainTenant
    The primary tenant (e.g. your organization's onmicrosoft.com domain) that has the delegated
    administration access to other tenants.

.PARAMETER ExceptionGroupName
    The display name of the group whose members will be excluded from being disabled even if they are
    otherwise deemed inactive.

.EXAMPLE
    # Example usage:
    .\DisableInactiveUsers.ps1 -ApplicationID "00000000-0000-0000-0000-000000000000" `
                              -ApplicationSecret "YOUR_SECRET_VALUE" `
                              -MainTenant "mytenant.onmicrosoft.com" `
                              -ExceptionGroupName "YourExceptionGroupName"
#>

###############################################################################
# 1. DEFINE ALL VARIABLES (PARAMETERS) AT THE TOP
###############################################################################
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$ApplicationID = "REPLACE_WITH_APP_ID",

    [Parameter(Mandatory=$true)]
    [string]$ApplicationSecret = "REPLACE_WITH_APP_SECRET",

    [Parameter(Mandatory=$true)]
    [string]$MainTenant = "mytenant.onmicrosoft.com",

    [Parameter(Mandatory=$true)]
    [string]$ExceptionGroupName = "YourExceptionGroupName"
)

# Set verbose output (optional: set to 'SilentlyContinue' to hide verbose messages)
$VerbosePreference = "Continue"

# Make these parameters accessible as global variables if needed within functions
$Global:ApplicationId      = $ApplicationID
$Global:ApplicationSecret  = $ApplicationSecret
$Global:MainTenant         = $MainTenant
$Global:ExceptionGroupName = $ExceptionGroupName

# Initialize global log
$global:logEntries = @()

Write-Host "`n=== Starting Inactive User Disabling Script ===`n" -ForegroundColor Green

###############################################################################
# 2. HELPER FUNCTIONS
###############################################################################

function Get-GraphToken {
<#
.SYNOPSIS
    Obtain an OAuth2 token for Microsoft Graph using client credentials.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$TenantID,

        [Parameter()]
        [string]$Scope = 'https://graph.microsoft.com/.default',

        [Parameter(Mandatory=$true)]
        [string]$ClientID,

        [Parameter(Mandatory=$true)]
        [string]$ClientSecret
    )

    $AuthBody = @{
        client_id     = $ClientID
        client_secret = $ClientSecret
        scope         = $Scope
        grant_type    = "client_credentials"
    }

    Write-Verbose "Requesting access token from Azure AD for TenantID: $TenantID"
    try {
        $AccessToken = Invoke-RestMethod `
            -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" `
            -Body $AuthBody -ErrorAction Stop

        Write-Verbose "Successfully obtained token for TenantID: $TenantID"
        return $AccessToken
    } catch {
        Write-Warning "Error obtaining access token for TenantID $TenantID: $($_.Exception.Message)"
        return $null
    }
}

function Connect-GraphAPI {
<#
.SYNOPSIS
    Connects to Microsoft Graph for a specific tenant using client credentials.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$ApplicationId,

        [Parameter(Mandatory=$true)]
        [string]$ApplicationSecret,

        [Parameter(Mandatory=$true)]
        [string]$TenantID
    )

    Write-Verbose "Removing old token if it exists..."
    $Script:GraphHeader = $null
    Write-Verbose "Attempting to log into Graph API for tenant: $TenantID"

    $AccessToken = Get-GraphToken -TenantID $TenantID -ClientID $ApplicationId -ClientSecret $ApplicationSecret
    if ($AccessToken) {
        $Script:GraphHeader = @{
            Authorization  = "Bearer $($AccessToken.access_token)"
            'Content-Type' = 'application/json'
        }
        Write-Verbose "Connection to $TenantID established."
        return $true
    } else {
        Write-Host "Could not log into the Graph API for tenant $TenantID" -ForegroundColor Red
        return $false
    }
}

function Log-Entry {
<#
.SYNOPSIS
    Logs a message to both the console and the global log array.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$message
    )

    Write-Output $message
    $global:logEntries += $message
}

function Disable-InactiveUsers {
<#
.SYNOPSIS
    Disables inactive users within a tenant, excluding members of a specified exception group.

.DESCRIPTION
    - Fetches all users (or only members if a certain domain is matched).
    - Determines if they are inactive based on last sign-in date.
    - Disables each inactive user except those in the exception group or meeting exclusion criteria.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$exceptionGroupId,

        [Parameter(Mandatory=$true)]
        [string]$TenantName
    )

    $VerbosePreference = "Continue"
    Write-Verbose "Starting Disable-InactiveUsers function for tenant: $TenantName"

    try {
        # Define inactivity thresholds
        $inactiveDate = (Get-Date).AddDays(-90)
        $accountAge   = (Get-Date).AddDays(-30)
        Write-Verbose "Inactivity threshold: $inactiveDate"
        Write-Verbose "Account age threshold: $accountAge"

        # Conditionally set API endpoint (example: exclude guests for a specific domain)
        if ($TenantName -eq "mytenant.com") {
            Write-Verbose "Tenant matches 'mytenant.com'. Excluding Guest users (only userType eq 'Member')."
            $Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$filter=accountEnabled eq true and userType eq 'Member'&`$select=id,displayName,userType,signInActivity,createdDateTime,onPremisesSyncEnabled&`$top=999"
        } else {
            # Default: fetch all enabled users
            $Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$filter=accountEnabled eq true&`$select=id,displayName,userType,signInActivity,createdDateTime,onPremisesSyncEnabled&`$top=999"
        }

        $Users = @()
        Write-Verbose "Fetching users from: $Uri"

        # Retrieve users via pagination
        do {
            try {
                $Data = Invoke-RestMethod -Uri $Uri -Headers $Script:GraphHeader -ErrorAction Stop
                $Users += $Data.value
                Write-Verbose "Fetched $($Data.value.Count) users. Total so far: $($Users.Count)"
                $Uri = $Data.'@odata.nextLink'
            } catch {
                Write-Warning "Error fetching data from Graph API: $_"
                return
            }
        } while ($Uri)

        Write-Verbose "Total users fetched: $($Users.Count) for tenant $TenantName"
        if (-not $Users) {
            Write-Host "No users returned. Exiting." -ForegroundColor Yellow
            return
        }

        # Build the final output with user details
        $FinalOutput = foreach ($User in $Users) {
            Write-Verbose "Processing user: $($User.displayName) with ID: $($User.id)"
            try {
                $GroupMembership = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)/memberOf" `
                    -Headers $Script:GraphHeader -ErrorAction Stop).value
                Write-Verbose "Group membership count for user $($User.displayName): $($GroupMembership.Count)"
            } catch {
                Write-Warning "Error fetching group membership for user $($User.displayName): $_"
                continue
            }

            # Check if the user is a member of the exception group
            $IsMemberOfGroup = $GroupMembership | Where-Object { $_.id -eq $exceptionGroupId }

            # Extract sign-in info
            $SignInActivity = $User.signInActivity
            $LastSuccessfulSigninDate = $SignInActivity.lastSuccessfulSignInDateTime

            $InactiveUserDays = if ($LastSuccessfulSigninDate) {
                (New-TimeSpan -Start $LastSuccessfulSigninDate).Days
            } else {
                "-"
            }

            Write-Verbose ("User {0} - Last Sign-In: {1}, Inactive Days: {2}" -f `
                $User.displayName, $LastSuccessfulSigninDate, $InactiveUserDays)

            # Output record
            [PSCustomObject]@{
                UserId                   = $User.id
                DisplayName              = $User.displayName
                CreatedDate              = $User.createdDateTime
                LastSuccessfulSigninDate = $LastSuccessfulSigninDate
                InactiveUserDays         = $InactiveUserDays
                IsMemberOfGroup          = $IsMemberOfGroup
                OnPremisesSyncEnabled    = $User.onPremisesSyncEnabled
                UserType                 = $User.userType
            }
        }

        if (-not $FinalOutput) {
            Write-Host "No users processed. Exiting script." -ForegroundColor Yellow
            return
        } else {
            Write-Verbose "Processed $($FinalOutput.Count) users. Proceeding to filtering."
        }

        # Filter users deemed inactive
        $InactiveUsers = $FinalOutput | Where-Object {
            ($_.LastSuccessfulSigninDate -lt $inactiveDate) -and
            ($_.CreatedDate -lt $accountAge) -and
            (-not $_.IsMemberOfGroup) -and
            ($null -eq $_.OnPremisesSyncEnabled -or $_.OnPremisesSyncEnabled -eq $false)
        }

        Write-Host "Found $($InactiveUsers.Count) inactive users to be disabled in tenant $TenantName." -ForegroundColor Cyan
        if ($InactiveUsers) {
            $InactiveUsers | Format-Table -AutoSize
        }

        # Disable inactive users
        foreach ($User in $InactiveUsers) {
            try {
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($User.UserId)" `
                                  -Method PATCH `
                                  -Headers $Script:GraphHeader `
                                  -Body (@{ accountEnabled = $false } | ConvertTo-Json) `
                                  -ErrorAction Stop

                $logEntry = "User: $($User.DisplayName) from tenant $TenantName was disabled. Last Sign-In: $($User.LastSuccessfulSigninDate), Inactive for: $($User.InactiveUserDays) days."
                Log-Entry -message $logEntry
                Write-Verbose $logEntry
            } catch {
                Write-Warning "Error disabling user $($User.DisplayName): $_"
            }
        }
    } catch {
        Write-Warning "Error in Disable-InactiveUsers function: $_"
    }
}

function Send-EmailReport {
<#
.SYNOPSIS
    Sends an email containing the global log of actions.

.DESCRIPTION
    Uses the MS Graph API endpoint for sending mail from a specified mailbox.
    Requires a valid token for the main tenant or a mailbox in that tenant.
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$recipient,

        [Parameter(Mandatory=$true)]
        [string[]]$logEntries
    )

    Write-Verbose "Obtaining token for sending email..."
    $accessToken = Get-GraphToken -TenantID $Global:MainTenant -ClientID $Global:ApplicationId -ClientSecret $Global:ApplicationSecret
    if (-not $accessToken) {
        Write-Warning "Failed to get token for sending email. No email will be sent."
        return
    }

    $headers = @{
        "Authorization" = "Bearer $($accessToken.access_token)"
        "Content-Type"  = "application/json"
    }

    $subject = "Inactive Users Disabled Report"
    $body    = $logEntries -join "`n"

    # Build the JSON payload
    $message = @{
        message = @{
            subject      = $subject
            body         = @{
                contentType = "Text"
                content     = $body
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $recipient
                    }
                }
            )
        }
    }

    $jsonMessage = $message | ConvertTo-Json -Depth 4

    # Adjust the 'sendMail' endpoint mailbox as needed
    $sendUrl = "https://graph.microsoft.com/v1.0/users/someone@mytenant.onmicrosoft.com/sendMail"
    Write-Verbose "Sending email to $recipient..."

    try {
        Invoke-RestMethod -Uri $sendUrl -Method Post -Headers $headers -Body $jsonMessage
        Write-Host "Email sent successfully to $recipient."
    } catch {
        Write-Warning "Failed to send email to $recipient: $_"
    }
}

###############################################################################
# 3. MAIN EXECUTION LOGIC
###############################################################################
function Main {
    Write-Host "Connecting to main tenant: $Global:MainTenant" -ForegroundColor Green
    $connected = Connect-GraphAPI -ApplicationId $Global:ApplicationId -ApplicationSecret $Global:ApplicationSecret -TenantID $Global:MainTenant
    if (-not $connected) {
        Write-Host "Connection to main tenant $($Global:MainTenant) failed. Exiting." -ForegroundColor Red
        return
    }

    Write-Host "Retrieving all delegated tenants for $($Global:MainTenant)..." -ForegroundColor Green
    try {
        $Tenants = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/contracts?`$top=999" -Method GET -Headers $Script:GraphHeader).value
    } catch {
        Write-Warning "Unable to retrieve tenants. Error: $($_.Exception.Message)"
        return
    }

    if ($Tenants) {
        $TenantCount     = $Tenants.Count
        Write-Host "$TenantCount tenants found. Attempting to process each..."
        $IncrementAmount = if ($TenantCount -eq 0) { 0 } else { 100 / $TenantCount }
        $i               = 0
        $ErrorCount      = 0

        foreach ($Tenant in $Tenants) {
            Write-Progress -Activity "Checking Tenant - Client Credentials" `
                           -Status "Progress -> Checking $($Tenant.defaultDomainName)" `
                           -PercentComplete $i `
                           -CurrentOperation "TenantLoop"

            Write-Verbose "Starting processing for tenant: $($Tenant.defaultDomainName)"
            $i = $i + $IncrementAmount

            # Connect to each sub-tenant
            $subConnected = Connect-GraphAPI -ApplicationId $Global:ApplicationId -ApplicationSecret $Global:ApplicationSecret -TenantID $Tenant.customerid
            if (-not $subConnected) {
                Write-Warning "Could not connect to tenant $($Tenant.defaultDomainName). Skipping..."
                $ErrorCount++
                continue
            }

            # Retrieve the exception group ID in this tenant
            try {
                $exceptionGroupId = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Method GET -Headers $Script:GraphHeader -ErrorAction Stop).value |
                    Where-Object { $_.displayName -eq $Global:ExceptionGroupName } |
                    Select-Object -ExpandProperty id
            } catch {
                Write-Warning "Error retrieving exception group in $($Tenant.defaultDomainName). $_"
                $ErrorCount++
                continue
            }

            if (-not $exceptionGroupId) {
                Write-Verbose "Exception group '$($Global:ExceptionGroupName)' not found in $($Tenant.defaultDomainName)."
            }

            # Disable inactive users
            Disable-InactiveUsers -exceptionGroupId $exceptionGroupId -TenantName $Tenant.defaultDomainName
        }

        Write-Host "Processing complete. Sending email reports..."

        # Send logs to one or more recipients
        Send-EmailReport -recipient "someone@mydomain.com" -logEntries $global:logEntries
        Send-EmailReport -recipient "anotherperson@mydomain.com" -logEntries $global:logEntries

        Write-Host "Script completed. Check your inbox(es) for the report." -ForegroundColor Green
    }
    else {
        Write-Warning "No delegated tenants found. Nothing to process."
    }
}

###############################################################################
# 4. RUN MAIN FUNCTION
###############################################################################
Main
Write-Host "`n=== Script Execution Complete ===`n" -ForegroundColor Green
