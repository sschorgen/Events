########################################################################################################################
#
#                                           DEMO SCRIPT FOR AOS SINGAPORE 2019
#       Script : New-MicrosoftTeams.ps1
#       Description : Automate Microsoft Teams Creation
#       Author : Sylver SCHORGEN
#
########################################################################################################################


# Import needed modules
# Modules must be imported in your environment
Import-Module MicrosoftTeams
Import-Module SharePointPnPPowerShellOnline

# Office 365 Admin user registered in a credential shared ressource named O365AdminAccount
# Automation variabes must be created as shared resources
$AzureAutomationAccount = Get-AutomationVariable -Name 'O365AdminAccount'
$O365Credentials = Get-AutomationPSCredential -Name $AzureAutomationAccount
$DefaultOwner = Get-AutomationVariable -Name 'DefaultTeamsOwner'

# SPSiteURL and SPTeamsListName are saved in a automation variables
# Automation variabes must be created as shared resources
$SPSiteUrl = Get-AutomationVariable -Name 'SharePointSite'
$SPTeamsListName = Get-AutomationVariable -Name 'SharePointTeamsList'

# Email information in order to send a recap email
# Automation variabes must be created as shared resources
$EmailFromAddress = Get-AutomationVariable -Name 'EmailAddressFrom'
$EmailToAddress = Get-AutomationVariable -Name 'EmailAddressTo'
$EmailSMTPServer = Get-AutomationVariable -Name 'EmailSMTPServer'
$Encoding = New-Object System.Text.utf8encoding

Write-Output "Variables configured"

# Connect to Microsoft Teams and SharePoint Online (with the PnP module)
# Module must be imported in your environment
Connect-MicrosoftTeams -Credential $O365Credentials
Connect-PnPOnline -Url $SPSiteUrl -Credential $O365Credentials
Write-Output "Connected to Microsof Teams and SharePoint Online"

# Getting all users
$Teams = Get-PnPListItem -List $SPTeamsListName
Write-Output "All Teams from the SharePoint list had been downloaded"

# Get every fields for each user in ordre to create the account
foreach ($Team in $Teams) {

    # If the team is set to be created (site column "Create" with value equals true)
    if ($Team.FieldValues.Create -eq $true) {
        $TeamName = "TO BE DEFINED"

        if($Team.FieldValues.TeamsType -eq "Company Department") {
            $TeamName = "DEPT - " + $Team.FieldValues.Title
        }
        elseif($Team.FieldValues.TeamsType -eq "Internal Project") {
            $TeamName = "PRJ INT - " + $Team.FieldValues.Title
        }
        elseif($Team.FieldValues.TeamsType -eq "External Project") {
            $TeamName = "PRJ EXT - " + $Team.FieldValues.Title
        }
        elseif($Team.FieldValues.TeamsType -eq "Fun") {
            $TeamName = "FUN - " + $Team.FieldValues.Title
        }
        
        # Checking if the team already exists or not
        $TeamExist = Get-Team -DisplayName $TeamName

        # If the teams does not exist, we create it and set the owners and members
        if($null -eq $TeamExist) {
            Write-Output "The Team $TeamName does not exist, let's create it !"
            $TeamDescription = $Team.FieldValues.Description
            $TeamsOwner = $Team.FieldValues.TeamsOwner.Email
            $TeamsMember = $Team.FieldValues.TeamsMember.Email
            $TeamToCreate = New-Team -DisplayName $TeamName -Description $TeamDescription -Owner $DefaultOwner

            Write-Output "The team $TeamName has been created."
            
            # Adding owners to the teams
            foreach ($Owner in $TeamsOwner) {
                $TeamToCreate | Add-TeamUser -User $Owner -Role Owner
                Write-Output "Owner $Owner has been added to the team"
            }
            # Adding the members to the teams
            foreach ($Member in $TeamsMember) {
                $TeamToCreate | Add-TeamUser -User $Member -Role Member
                Write-Output "Member $Member has been added to the team"
            }

            # Creating the email title and body
            $EmailSubject = "Office 365 - Team $TeamName has been created"
            $EmailBody = "Hi, `n `n"
            $EmailBody += "The Team $TeamName has been created. `n"
            $EmailBody += "Owners are $TeamsOwner. `n"
            $EmailBody += "Members are $TeamsMember. `n"
            $EmailBody += "Best regards, `n"
            $EmailBody += "The Office 365 Automation Administrator"

            # Sending the email
            Send-MailMessage -Credential $O365Credentials -From $EmailFromAddress -To $EmailToAddress -Subject $EmailSubject -Body $EmailBody -SmtpServer $EmailSMTPServer -UseSSL -Encoding $Encoding
            Write-Output "An email has been send to $EmailToAddress"

            # Setting the field "user to create" to false
            Set-PnPListItem -List $SPTeamsListName -Identity $Team.Id -Values @{"Create" = $false}
        }
    }
    
}