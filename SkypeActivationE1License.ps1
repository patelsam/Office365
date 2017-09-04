# The script was used for TransportNSW SfB rollout project
#Established connection to O365 tenant
Write-Host "***********************************************"
Write-Host "Connect to Office365 Tenant" -ForegroundColor Cyan
Write-Host "***********************************************"

Import-Module MSOnline
$O365Cred = Get-Credential
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $O365Session
Connect-MsolService –Credential $O365Cred

#Start Logging..
$VerbosePreference = "Continue"
$LogPath = Split-Path $MyInvocation.MyCommand.Path
#Get-ChildItem "$LogPath\*.log" | Where LastWriteTime -LT (Get-Date).AddDays(-15) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'dd-MM-yyyy_HH-mmtt').log"
Start-Transcript $LogPathName -Append

#Read data from churn sheet
Write-Host "***********************************************"
Write-Host "Reading data from csv file...." -ForegroundColor Cyan
Write-Host "***********************************************"
$csvdata = Import-Csv "C:\Scripts\SkypeActivation\COE_churnsheet_UnLicensed.csv"

#Create Custom Lincese for Skype only
$sku = "TransportCloud:STANDARDPACK"
$disabled ='FLOW_O365_P1','Deskless','POWERAPPS_O365_P1','TEAMS1','PROJECTWORKMANAGEMENT','SWAY','INTUNE_O365','YAMMER_ENTERPRISE','SHAREPOINTWAC','EXCHANGE_S_STANDARD','SHAREPOINTSTANDARD'
$myO365Sku1 = New-MsolLicenseOptions -AccountSkuId $sku -DisabledPlans $disabled;

foreach($user in $csvdata)
{
#Set Usage location and Assign Customised License to each user
Write-Host "***********************************************"
Write-Host "Setting UsageLocation and Assigning License.." -ForegroundColor Cyan
Write-Host "***********************************************"
Get-MsolUser -UserPrincipalName $user.email | Set-MsolUser -UsageLocation AU
Get-MsolUser -UserPrincipalName $user.email | Set-MsolUserLicense -AddLicenses $sku -LicenseOptions $myO365Sku1

#Get Objectid for each user and add it to cloud-only security groups
Write-Host "***********************************************"
Write-Host "Adding user to cloud-only AV Groups.." -ForegroundColor Cyan
Write-Host "***********************************************"

$userid = Get-MsolUser -UserPrincipalName $user.email | Select objectid

$group = get-msolgroup -all | where {$_.Displayname -eq  $user.cloudgroup} | Select ObjectId
add-msolgroupmember -groupobjectid $group.objectid -groupmembertype “user” -GroupMemberObjectId $userid.ObjectID

}


Remove-PSSession $O365Session
Stop-Transcript
