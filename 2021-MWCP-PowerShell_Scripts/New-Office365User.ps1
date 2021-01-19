# Import needed modules
Import-Module SharePointPnPPowerShellOnline
Import-Module AzureAD

# Office 365 Admin user registered in a credential shared ressource named O365-Admin
$AzureAutomationAccount = "O365-Admin"
$O365Credentials = Get-AutomationPSCredential -Name $AzureAutomationAccount

# SPSiteURL, SPListName, SkuID and temp password saved in an automation variable
$SPSiteUrl = Get-AutomationVariable -Name 'SharePointUserSite'
$SPListName = Get-AutomationVariable -Name 'SharePointUserList'
$SkuID = Get-AutomationVariable -Name 'AzureSkuID'
$TempUserPassword = Get-AutomationVariable -Name 'TempUserPassword'

# Generation of a generic password for the Office 365 account
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $TempUserPassword
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
        $UserMail = "$UserFirstname.$UserLastname@schorgenlab.com"
        $UserCountry = "AU"
        $Country = "Australia"
        $City = "Sydney"

        # User creation
        $UserToCreate = New-AzureADUser -GivenName $UserFirstname -Surname $UserLastname -DisplayName "$UserFirstname $UserLastname" -UserPrincipalName $UserMail -MailNickName "$UserFirstname.$UserLastname" -AccountEnabled $true `
        -PasswordProfile $PasswordProfile -JobTitle $UserJobTitle -Department $UserDepartment -UsageLocation $UserCountry -Country $Country -City $City -TelephoneNumber $PhoneNumber
        
        # Assigning the manager to the newly created user
        Set-AzureADUserManager -ObjectId $UserToCreate.ObjectId -RefObjectId $UserManager.ObjectId

        # Assigning the license to the newly created user
        Set-AzureADUserLicense -ObjectId $UserToCreate.ObjectId -AssignedLicenses $Licenses

        Write-Output "The user $UserFirstname $UserLastname had been created"

        # Setting the field "user to create" to false
        Set-PnPListItem -List $SPListName -Identity $User.Id -Values @{"Create" = $false}        
    }
    
}