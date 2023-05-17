# PowerShell samples

# Create certificate for login
# stolen from https://www.powershellgallery.com/packages/ExchangeOnlineManagement/1.0.1/Content/Create-SelfSignedCertificate.ps1 and other places
# Google "Create-SelfSignedCertificate.ps1" for more places

.\Create-SelfSignedCertificate.ps1 -CommonName "Techorama" -StartDate 2023-05-16 -EndDate 2025-05-16 -Password (ConvertTo-SecureString "pass@word1" -Force -AsPlainText)

# Run in an Admin console
Install-Module PnP.PowerShell -Force -Scope AllUsers
Install-Module Microsoft.Online.SharePoint.PowerShell -Force -Scope AllUsers
Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope AllUsers
Install-Module Microsoft.PowerApps.PowerShell -AllowClobber -Scope AllUsers

# Only gets the modules installed in the current PowerShell host
Get-InstalledModule

# Gets all of the modules installed on the machine
Get-Module -ListAvailable

# Make sure a module is installed where both PowerShells and all users can use it
Get-InstalledModule -Name Microsoft.Online.SharePoint.PowerShell | Select-Object InstalledLocation

Update-Module Microsoft.Online.SharePoint.PowerShell
  
# Use stored credentials
Add-PnPStoredCredential -Name "https://tenant.sharepoint.com" -Username yourname@tenant.onmicrosoft.com
Connect-PnPOnline -Url "https://tenant.sharepoint.com”
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/hr”

# Grab stored crenetials for use in other modules
$Credentials = Get-PnPStoredCredential -Name "https://tenant.sharepoint.com" 
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credentials $Credentials

# Upload a file
$web = https://tenant.sharepoint.com/sites/hr
$folder = "Shared Documents"
Connect-PnPOnline -Url $web
Add-PnPFile -Path '.\Boot fairs with Graphic design.docx' -Folder $folder

# Add a folder
Add-PnPFolder -Name "Folder 1" -Folder $folder
Add-PnPFile -Path '.\Building materials licences to budget for Storytelling.docx'  -Folder "$folder\Folder 1"

# Get File sharing information
$docliblist = Get-PnPList -Includes DefaultViewUrl,IsSystemList | Where-Object -Property IsSystemList -EQ -Value $false | Where-Object -Property BaseType -EQ -Value "DocumentLibrary"

    Foreach ($doclib in $docliblist) 
        {
	        $doclibTitle = $doclib.Title
            $docs = Get-PnPListItem -List $DocLib
	        $docs | ForEach-Object { Get-PnPProperty -ClientObject $_ -Property HasUniqueRoleAssignments | Out-Null}
            foreach ($doc in $docs) {
                [pscustomobject]@{
                    Library = $doclibTitle
                    Filename = $doc.FieldValues.FileLeafRef
                    Shared = $doc.HasUniqueRoleAssignments
                    }
                }
            }
        


# Bulk Undelete Files example
# Actually does work
Connect-PnPOnline -Url https://sadtenant.sharepoint.com/ -Credentials SadTenantAdmin

# Get the files to restore
$bin = Get-PnPRecycleBinItem | Where-Object -Property Leafname -Like -Value "*.jpg"  | Where-Object -Property Dirname -Like -Value "Important Photos/Shared Documents/*"  | Where-Object -Property DeletedByEmail -EQ -Value baduser@sadtenant.phooey

# Show how many files we're going to restore
$bin.count

# Restore the files
$bin | ForEach-Object -begin { $a = 0} -Process {Write-Host "$a - $($_.LeafName)" ; $_ | Restore-PnPRecycleBinItem -Force ; $a++ } -End { Get-Date }

# restore a subset of files
($bin[20001..30000]) | ForEach-Object -begin { $a = 0} -Process {Write-Host "$a - $($_.LeafName)" ; $_ | Restore-PnPRecycleBinItem -Force ; $a++ } -End { Get-Date }

# https://www.toddklindt.com/PoshRestoreSPOFiles

# Create sites examples
# No Group, No Team, No Bueno Can be Groupified later
New-SPOSite 

# Group, No Team Can be Teamified later
New-PnPSite -Type TeamSite -Title "Modern Team Site" -Alias ModernTeamSite -IsPublic

# There is no later!
New-Team -DisplayName "Fancy Group" -Description "Fancy Group made by PowerShell?" -Alias FancyGroup -AccessType Public

# Save a site as a template
Get-PnPSiteTemplate -Out customer.xml
Add-PnPListFoldersToSiteTemplate -Path customer.xml -List 'Data Storage' -Recursive
Invoke-PnPSiteTemplate -Path customer.xml -Handlers Lists, SiteSecurity

# See other template commands
Get-Command -Module PnP.PowerShell -Name *temp*






# Group Membership example
# Set some values 
# Name of Unified Group whose owners and membership we want to copy 
$source = "Regulations"
# Name of Unified Group whose owners and membership we want to populate 
$destination = "Empty"
# Whether to overwrite Destination membership or merge them 
$mergeusers = $false
# Check to see if PnP Module is loaded 
$pnploaded = Get-Module PnP.PowerShell
if ($pnploaded -eq $false) {
    	Write-Host "Please load the PnP PowerShell and run again" 
    	Write-Host "install-module PnP.PowerShell" 
    	break 
    } 

# PnP Module is loaded
# Check to see if user is connected to Microsoft Graph 
try 
{ 
    $owners = Get-PnPMicrosoft365GroupOwner -Identity $source 
} 
catch [System.InvalidOperationException] 
{ 
    Write-Host "No connection to Microsoft Graph found"  -BackgroundColor Black -ForegroundColor Red 
    Write-Host "No Azure AD connection, please connect first with Connect-PnPOnline -Graph" -BackgroundColor Black -ForegroundColor Red 
break 
} 
catch [System.ArgumentNullException] 
{ 
        Write-Host "Group not found"  -BackgroundColor Black -ForegroundColor Red 
        Write-Host "Verify connection to Azure AD with Connect-PnPOnline -Graph" -BackgroundColor Black -ForegroundColor Red 
        Write-Host "Use Get-PnPUnifiedGroup to get Unified Group names"  -BackgroundColor Black -ForegroundColor Red 
        break 
} 
catch 
{ 
    Write-Host "Some other error"   -BackgroundColor Black -ForegroundColor Red 
break 
}

$members = Get-PnPMicrosoft365GroupMember -Identity $source
if ($mergeusers -eq $true) { 
     # Get existing owners and members of Destination so that we can combine them 
    $ownersDest = Get-PnPMicrosoft365GroupOwner -Identity $destination 
    $membersDest = Get-PnPMicrosoft365GroupMember -Identity $destination
    # Add the two lists together so we don't overwrite any existing owners or members in Destination 
    $owners = $owners + $ownersDest 
    $members = $members + $membersDest 
    }
# Set the owners and members of Destination 
$owners | ForEach-Object -begin  {$ownerlist  = @() } -process {$ownerlist += $($_.UserPrincipalName) } 
$members | ForEach-Object -begin  {$memberlist  = @() } -process {$memberlist += $($_.UserPrincipalName) }
Set-PnPMicrosoftGroup -Identity $destination -Members $memberlist -Owners $ownerlist
# https://www.toddklindt.com/PoshCopyO365GroupMembers

# Get all the Flows

# Connect to PowerApps
Add-PowerAppsAccount

# Get all the Flows
# Uses the msonline module. Naughty, naughty
Get-AdminFlow | ForEach-Object { $ownername = (Get-MsolUser -ObjectId $_.CreatedBy.userId).DisplayName ; $owneremail = (Get-MsolUser -ObjectId $_.CreatedBy.userId).UserPrincipalName ; Write-Host $_.DisplayName, $ownername, $owneremail }

# Get your own Flows
Get-Flow

# Get the PowerApps
Get-PowerApp

# Disable Flow Button
# SPO Method
Connect-SPOService -Url https://flowhater-admin.sharepoint.com
$val = [Microsoft.Online.SharePoint.TenantAdministration.FlowsPolicy]::Disabled 
Set-SPOSite -Identity https://flowhater.sharepoint.com/sites/SadSite -DisableFlows $val

# PnP Method
Connect-PnPOnline -Url https://flowhater.sharepoint.com/sites/SadSite
Set-PnPSite -DisableFlows:$true


