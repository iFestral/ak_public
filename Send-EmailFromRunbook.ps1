<#
.SYNOPSIS
  This runbook sends an email using Microsoft Graph from Azure Automation.

.DESCRIPTION
  The runbook connects to Azure using a managed identity, authenticates with Microsoft Graph, and sends an email to the specified recipient.

.EXAMPLE
  Update variables in section # Define email details
  
  # Run the script to send a test email
  .\Send-EmailFromRunbook.ps1
#>

# Define email details
$emailSubject = "Test Azure Runbook"
$emailBody = "This is a test email sent from Azure Automation Runbook."
$recipientEmail = ""
$userId = ""

# Function to log messages
Function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Info', 'Error')]
        [string]$MessageType
    )

    begin {
        $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    process {
        $msg = "[$MessageType] [$date] $Message"
    }
    end {
        Write-Output $msg
    }
}

# Function to send an email via Microsoft Graph
Function Send-Email {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$BodyContent,

        [Parameter(Mandatory = $true)]
        [string]$ToAddress,

        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        $params = @{
            Message = @{
                Subject      = $Subject
                Body         = @{
                    ContentType = "Text"
                    Content     = $BodyContent
                }
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $ToAddress
                        }
                    }
                )
            }
        }

        Send-MgUserMail -UserId $UserId -BodyParameter $params
        Write-Log -Message "Email sent successfully to $ToAddress." -MessageType 'Info'
    }
    catch {
        Write-Log -Message "Failed to send email. Error: $_" -MessageType 'Error'
        throw $_
    }
}

# Main script logic
try {
    # Ensure the script is authenticated with Azure
    if (!$az_connection) {
        Write-Log -Message "Connecting to Azure..." -MessageType 'Info'
        $az_connection = Connect-AzAccount -Identity
        Write-Log -Message "Connected to Azure successfully." -MessageType 'Info'
    }

    if (!$msgraph_connection) {
        Write-Log -Message "Connecting to Microsoft Graph..." -MessageType 'Info'
        $AccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
        $AccessTokenSecure = ConvertTo-SecureString -String $AccessToken.Token -AsPlainText -Force
        $msgraph_connection = Connect-MgGraph -AccessToken $AccessTokenSecure
        Write-Log -Message "Connected to Microsoft Graph successfully." -MessageType 'Info'
    }

    # Send the email
    Write-Log -Message "Sending an email to $recipientEmail..." -MessageType 'Info'
    Send-Email -Subject $emailSubject -BodyContent $emailBody -ToAddress $recipientEmail -UserId $userId

} catch {
    Write-Log -Message "Runbook failed. Error: $_" -MessageType 'Error'
    throw $_
}

