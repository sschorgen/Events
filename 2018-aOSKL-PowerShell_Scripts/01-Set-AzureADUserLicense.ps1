########################################################################################################################
#
#                                           DEMO SCRIPT FOR AOS KUALA LUMPUR
#       Script : Set-AzureADUserLicense
#       Description : Automate O365 license management
#       Author : Sylver SCHORGEN
#
########################################################################################################################


# Import needed modules
Import-Module AzureAD

# Office 365 Admin user registered in a credential shared ressource named O365-Admin
# Automation variabes must be created as shared resources
$AzureAutomationAccount = Get-AutomationVariable -Name 'O365AdminAccount'
$O365Credentials = Get-AutomationPSCredential -Name $AzureAutomationAccount

# SkuID saved in an encryted variable
# Automation variabes must be created as shared resources
$SkuID = Get-AutomationVariable -Name 'SkuID'

# Email information in order to send a recap email
# Automation variabes must be created as shared resources
$EmailFromAddress = Get-AutomationVariable -Name 'EmailAddressFrom'
$EmailToAddress = Get-AutomationVariable -Name 'EmailAddressTo'
$EmailSMTPServer = Get-AutomationVariable -Name 'EmailSMTPServer'
$Encoding = New-Object System.Text.utf8encoding

# Configuring User Location- NC is for New Caledonia
# Automation variabes must be created as shared resources
$UsageLocation = Get-AutomationVariable -Name 'UsageLocation'

# SkuID information from my tenant
$Sku = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$Sku.SkuId = $SkuID

# Licenses Object to assign to users
$Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$Licenses.AddLicenses = $Sku

# Connect to AzureAD
Connect-AzureAD -Credential $O365Credentials

# Get all users in the E3 security group I created
# This group is mandatory in order to know to which I need to apply the license to
$UserToLicenseGroup = Get-AzureADGroupMember -ObjectId (Get-AzureADGroup | Where-Object DisplayName -eq "O365-E3").ObjectID

foreach ($User in $UserToLicenseGroup) {
    # Verify if the user already have a license
    $UserIsLicensed = $false

    foreach ($License in $User.AssignedLicenses) {
        if ($License.SkuID -eq $Sku.SkuId) {
            $UserIsLicensed = $true

            Write-Output "User $($User.DisplayName) already have a license !"
        }
    }

    # If the user doesn't have a license, we assign one and send an email to the administrator
    if ($UserIsLicensed -eq $false) {

        Write-Output "User $($User.DisplayName) does not have any license !"
        Write-Output "Assigning an O365 E3 licence to $($User.DisplayName) and setting up the location to $UsageLocation"
        Set-AzureADUser -ObjectId $User.ObjectID -UsageLocation $UsageLocation
        Set-AzureADUserLicense -ObjectId $User.ObjectID -AssignedLicenses $Licenses
        
        $EmailSubject = "Office 365 - User E3 License assigned for " + $User.DisplayName
        $EmailBody = "Hi, `n `n"
        $EmailBody += "An Office 365 E3 license has been assigned to $($User.DisplayName). `n `n"
        $EmailBody += "Best Regards, `n"
        $EmailBody += "Office 365 Automation Administrator"
        
        Write-Output "Sending email validation to $EmailToAddress"
        Send-MailMessage -Credential $O365Credentials -From $EmailFromAddress -To $EmailToAddress -Subject $EmailSubject -Body $EmailBody -SmtpServer $EmailSMTPServer -UseSSL -Encoding $Encoding
    }
}