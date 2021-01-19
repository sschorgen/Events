# Import needed modules
Import-Module AzureAD

# Office 365 Admin user registered in a credential shared ressource named O365-Admin
$AzureAutomationAccount = "O365-Admin"
$O365Credentials = Get-AutomationPSCredential -Name $AzureAutomationAccount

# SkuID saved in an encryted variable
$SkuID = Get-AutomationVariable -Name 'SkuID'

# Configuring User Location
$UsageLocation = Get-AutomationVariable -Name 'UsageLocation'

# SkuID information from my tenant
$Sku = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$Sku.SkuId = $SkuID

# Licenses Object to assign to users
$Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$Licenses.AddLicenses = $Sku

# Get security group containing users to license
$SecurityGroup = Get-AutomationVariable -Name 'SecurityGroup'
# Connect to AzureAD
Connect-AzureAD -Credential $O365Credentials

# Get all users in the E3 security group I created
# This group is mandatory in order to know to which I need to apply the license to
$UserToLicenseGroup = Get-AzureADGroupMember -ObjectId (Get-AzureADGroup | Where-Object DisplayName -eq $SecurityGroup).ObjectID

foreach ($User in $UserToLicenseGroup) {
    # Verify if the user already have a license
    $UserIsLicensed = $false

    foreach ($License in $User.AssignedLicenses) {
        if ($License.SkuID -eq $Sku.SkuId) {
            $UserIsLicensed = $true

            Write-Output "User $($User.DisplayName) already has a license !"
        }
    }

    # If the user doesn't have a license, we assign one and send an email to the administrator
    if ($UserIsLicensed -eq $false) {

        Write-Output "User $($User.DisplayName) does not have any license !"
        Write-Output "Assigning an O365 E3 licence to $($User.DisplayName) and setting up the location to $UsageLocation"
        Set-AzureADUser -ObjectId $User.ObjectID -UsageLocation $UsageLocation
        Set-AzureADUserLicense -ObjectId $User.ObjectID -AssignedLicenses $Licenses
    }
}