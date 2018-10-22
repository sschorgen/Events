########################################################################################################################
#
#                                           DEMO SCRIPT FOR AOS KUALA LUMPUR
#       Script : Set-ExchangeOnlineDistributionGroup
#       Description : Automate O365 license management
#       Author : Sylver SCHORGEN
#
########################################################################################################################


# Import needed modules
Import-Module SharePointPnPPowerShellOnline

# Office 365 Admin user registered in a credential shared ressource named O365-Admin
# Automation variabes must be created as shared resources
$AzureAutomationAccount = Get-AutomationVariable -Name 'O365AdminAccount'
$O365Credentials = Get-AutomationPSCredential -Name $AzureAutomationAccount

# SPSiteURL and SPListName saved in an automation variable
# Automation variabes must be created as shared resources
$SPSiteUrl = Get-AutomationVariable -Name 'SharePointSite'
$SPListName = Get-AutomationVariable -Name 'SharePointDLList'

# Email information in order to send a recap email
# Automation variabes must be created as shared resources
$EmailFromAddress = Get-AutomationVariable -Name 'EmailAddressFrom'
$EmailToAddress = Get-AutomationVariable -Name 'EmailAddressTo'
$EmailSMTPServer = Get-AutomationVariable -Name 'EmailSMTPServer'
$Encoding = New-Object System.Text.utf8encoding

Write-Output "Variables configured"

# Using SharePoint PnP to connect to SharePoint Online and connecting to Exchange Online
Connect-PnPOnline -Url $SPSiteUrl -Credentials $O365Credentials
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Credentials -Authentication Basic -AllowRedirection
Import-Module (Import-PSSession -Session $Session -AllowClobber -DisableNameChecking) -Global

Write-Output "Exchange Online module loaded and connected to SharePoint Online (PnP)"

$DistributionGroups = Get-PnPListItem -List $SPListName

Write-Output "Managed distribution groups downloaded from the SharePoint list"

# Loop through all the distribution groups
foreach ($DistributionGroup in $DistributionGroups) {

    # Get all the user from the distribution group
    $GroupMembers = Get-DistributionGroupMember -Identity $DistributionGroup.FieldValues.Title | Select-Object DisplayName, PrimarySMTPAddress

    # Loop through all the users in the distribution groups
    foreach ($User in $DistributionGroup) {

        # Loop through all the mail address already in the distribution group in order to add only the new mail addresses
        foreach($EmailAddress in $User.FieldValues.User.Email){
            
            $UserAlreadyMember = $false

            # Validate if the user is already in the distribution group or not
            foreach ($Member in $GroupMembers) {
                if ($Member.PrimarySMTPAddress -eq $EmailAddress) {
                    $UserAlreadyMember = $true
                }
            }

            # If the user is not in the distribution group, we add him and send an email to the administrator
            if ($UserAlreadyMember -eq $false) {
                Add-DistributionGroupMember -Identity $User.FieldValues.Title -Member $EmailAddress -ErrorAction SilentlyContinue
                
                $EmailSubject = "Office 365 - Distribution Group " + $User.FieldValues.Title + " modification"
                $EmailBody = "Hi, `n `n"
                $EmailBody += "The mail address $EmailAddress has been added to the distribution group " + $User.FieldValues.Title + ". `n `n"
                $EmailBody += "Best Regards, `n"
                $EmailBody += "Office 365 Automation Administrator"

                Write-Output "User $EmailAddress added to DG $($User.FieldValues.Title)"
                Send-MailMessage -Credential $O365Credentials -From $EmailFromAddress -To $EmailToAddress -Subject $EmailSubject -Body $EmailBody -SmtpServer $EmailSMTPServer -UseSSL -Encoding $Encoding
            } else {
                Write-Output "User $EmailAddress already in the DG $($User.FieldValues.Title)"
            }
        }
    }
}

# Removing the PowerShell session
Remove-PSSession $Session