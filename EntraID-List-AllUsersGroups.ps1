<#
.SYNOPSIS
This script collects EntraID users and their groups related to roles, then updates a SharePoint list with their information.

.DESCRIPTION
The script connects to Azure Key Vault to retrieve the username and password for the automation account. It then uses the obtained credentials to connect to SharePoint and retrieve a list of EntraID users. 
The script checks if each user is already present in the SharePoint list. If not, it adds a new list item with the user's information. 
If the user already exists in the list, it updates the corresponding list item with the user's group information.

.PARAMETER None
This script does not accept any parameters.

.EXAMPLE
.\EntraID-Lists-AllUsersGroups.ps1
Runs the script to collect EntraID users and update the SharePoint list.
#>

# Function block
Function Write-Log {
    
    [CmdletBinding()]
    param(

        [Parameter(Mandatory = $true)] $Message,
        [ValidateSet('Info', 'Error')] $MessageType
    )
    begin {  
        $date = Get-Date
    }
    process {
        if ($MessageType -eq 'Info') {

            $msg = "[INFO] [$date] $Message"

        }
        elseif ($MessageType -eq 'Error') {

            $msg = "[ERROR] [$date] $Message" 

        }

        Write-Output $msg
        
    }
}

# Variables
$sharepoint_site_url = "https://akdotms-my.sharepoint.com/personal/svc-powerautomate_akdotms_cloud/"
$sharepoint_list_name = "Users/Groups List"

# Azure Key Vault Variables
$kv_name = "kv-automation-ne"
$automation_name_secret = "automation-user"
$automation_password_secret = "automation-user-password"

# Connect Azure
if (!$az_connection) {
    $az_connection = Connect-AzAccount -Identity
}

# Get Automation username and password from KeyVault
$automation_name = Get-AzKeyVaultSecret -VaultName $kv_name -Name $automation_name_secret -AsPlainText
$automation_password = Get-AzKeyVaultSecret -VaultName $kv_name -Name $automation_password_secret -AsPlainText


# Prepare credentials for PnP connection
[securestring]$automation_password_ss = ConvertTo-SecureString $automation_password -AsPlainText -Force
[pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($automation_name, $automation_password_ss)

# Connect Graph API
if (!$msgraph_connection) {

    $AccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
    $AccessTokenSecure = ConvertTo-SecureString -String $AccessToken.Token -AsPlainText -Force
    $msgraph_connection = Connect-MgGraph -AccessToken $AccessTokenSecure

}

# Collecting EntraID Users
Write-Log -MessageType Info -Message "Processing EntraID Users..."

$Properties = @("AccountEnabled", "UserType", "Id", "DisplayName", "UserPrincipalName", "CompanyName", "JobTitle")
$users = Get-MgUser -All -ConsistencyLevel eventual -Filter "accountEnabled eq true and userType eq 'Member'" -Property $Properties | Where-Object { (($_.JobTitle -ne "Contractor") -and ($_.JobTitle -ne "Shared Mailbox")) }

# Connect PnP with credentials
Write-Log -MessageType Info -Message "Connecting to SharePoint List"

if (!$pnp_connection) {
    $pnp_connection = Connect-PnPOnline -Url $sharepoint_site_url -Credentials $credObject -WarningAction Ignore
}

#Get List of items to Update
$ListItems = Get-PnPListItem -List $sharepoint_list_name -ErrorAction Stop

foreach ($user in $users) {

    if (!$ListItems.fieldValues.Title.Contains($user.UserPrincipalName)) {
        try {
            [array]$groups = Get-MgUserMemberOf -UserId $user.Id | Where-Object { ($_.AdditionalProperties.displayName -like "akdotms.grp.user.*") }
            $Create_item = Add-PnPListItem -List $sharepoint_list_name -Values @{"User" = "$($user.UserPrincipalName)"; "Title" = "$($user.UserPrincipalName)"; "Groups" = $groups.AdditionalProperties.displayName; }
            Write-Log -MessageType Info -Message "Created for $($user.UserPrincipalName)"
        }
        catch {
            $msg = "Failed to create List item for $($user.UserPrincipalName): "
            $msg += $_.Exception.Message
            Write-Log -MessageType Error -Message $msg
        }    
    }
    else {
        try {
            [array]$groups = Get-MgUserMemberOf -UserId $user.Id | Where-Object { ($_.AdditionalProperties.displayName -like "akdotms.grp.user.*") }
            $ListItemUpdate = (Get-PnPListItem -List $sharepoint_list_name).fieldValues | Where-Object { $_.Title -eq $($user.UserPrincipalName) }
            $Update_item = Set-PnPListItem -List $sharepoint_list_name -Id $ListItemUpdate.ID -Values @{"Groups" = $groups.AdditionalProperties.displayName;}
            Write-Log -MessageType Info -Message "Updated $($user.UserPrincipalName)"
        }
        catch {
            $msg = "Failed to update List item for $($user.UserPrincipalName): "
            $msg += $_.Exception.Message
            Write-Log -MessageType Error -Message $msg
        }
    }
}
