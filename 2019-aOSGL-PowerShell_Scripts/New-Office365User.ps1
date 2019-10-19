########################################################################################################################
#
#                                           DEMO SCRIPT FOR AOS SINGAPORE 2019
#       Script : New-Office365User.ps1
#       Description : Automate O365 license management
#       Author : Sylver SCHORGEN
#
########################################################################################################################

# Import needed modules
Import-Module SharePointPnPPowerShellOnline
Import-Module AzureAD

# Office 365 Admin user registered in a credential shared ressource named O365-Admin
# Automation variabes must be created as shared resources
$AzureAutomationAccount = Get-AutomationVariable -Name 'O365AdminAccount'
$O365Credentials = Get-AutomationPSCredential -Name $AzureAutomationAccount

# SPSiteURL, SPListName and SkuID saved in an automation variable
# Automation variabes must be created as shared resources
$SPSiteUrl = Get-AutomationVariable -Name 'SharePointSite'
$SPListName = Get-AutomationVariable -Name 'SharePointUserList'
$SkuID = Get-AutomationVariable -Name 'SkuID'

# Email information in order to send a recap email
# Automation variabes must be created as shared resources
$EmailFromAddress = Get-AutomationVariable -Name 'EmailAddressFrom'
$EmailToAddress = Get-AutomationVariable -Name 'EmailAddressTo'
$EmailSMTPServer = Get-AutomationVariable -Name 'EmailSMTPServer'
$Encoding = New-Object System.Text.utf8encoding

# Generation of a generic password for the Office 365 account
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = '#2017P@$_w0rd1999!'
$PasswordProfile.ForceChangePasswordNextLogin = $true

# SkuID information from my tenant
$Sku = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$Sku.SkuId = $SkuID

# Licenses Object to assign to users
$Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$Licenses.AddLicenses = $Sku

Write-Output "Variables configured"

# Connect to AzureAD and SharePoint Online (with the PnP module)
Connect-PnPOnline -Url $SPSiteUrl -Credentials $O365Credentials
Connect-AzureAD -Credential $O365Credentials

Write-Output "Connected to SharePoint Online (PowerShell PnP) and Azure AD"

# Getting all users
$Users = Get-PnPListItem -List $SPListName
Write-Output "All users from the SharePoint list had been downloaded"

# Get every fields for each user in ordre to create the account
foreach ($User in $Users) {

    # If the user is set to be created (site column "Create" with value equals true)
    if ($User.FieldValues.Create -eq $true) {
        $UserFirstname = $User.FieldValues.Firstname
        $UserLastname = $User.FieldValues.Lastname
        $UserJobTitle = $User.FieldValues.JobTitle
        $UserDepartment = $User.FieldValues.Department
        $PhoneNumber = $User.FieldValues.Phone
        $UserManager = Get-AzureADUser -ObjectId $User.FieldValues.Manager.Email
        $UserMail = "$UserFirstname.$UserLastname@demoaossg.onmicrosoft.com"
        $UserCountry = "SG"
        $Country = "Singapore"
        $City = "Singapore"

        # User creation
        $UserToCreate = New-AzureADUser -GivenName $UserFirstname -Surname $UserLastname -DisplayName "$UserFirstname $UserLastname" -UserPrincipalName $UserMail -MailNickName "$UserFirstname.$UserLastname" -AccountEnabled $true `
        -PasswordProfile $PasswordProfile -JobTitle $UserJobTitle -Department $UserDepartment -UsageLocation $UserCountry -Country $Country -City $City -TelephoneNumber $PhoneNumber
        
        # Assigning the manager to the newly created user
        Set-AzureADUserManager -ObjectId $UserToCreate.ObjectId -RefObjectId $UserManager.ObjectId

        # Assigning the license to the newly created user
        Set-AzureADUserLicense -ObjectId $UserToCreate.ObjectId -AssignedLicenses $Licenses

        # Creating the email title and body
        $EmailSubject = "Office 365 - User $UserFirstname $UserLastname created"
        $EmailBody = "Hi, `n `n"
        $EmailBody += "The user $UserFirstname $UserLastname has been created. `n"
        $EmailBody += "An Office 365 E3 license has been assigned to the user.`n `n"
        $EmailBody += "Best regards, `n"
        $EmailBody += "The Office 365 Automation Administrator"

        # Sending the email
        Send-MailMessage -Credential $O365Credentials -From $EmailFromAddress -To $EmailToAddress -Subject $EmailSubject -Body $EmailBody -SmtpServer $EmailSMTPServer -UseSSL -Encoding $Encoding
        Write-Output "The user $UserFirstname $UserLastname had been created"
        Write-Output "An email has been send to $EmailToAddress"

        # Setting the field "user to create" to false
        Set-PnPListItem -List $SPListName -Identity $User.Id -Values @{"Create" = $false}        
    }
    
}