# Using Graph SDK

# Connect through the browser
# Make sure to have the right profile in focus
Connect-MgGraph

# Graph Context information
Get-MgContext

# Get the scopes
(Get-MgContext).Scopes

# Get the users
Get-MgUser

# Might need to specify the scope
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.Read.All"

# Play with the users
$MGUserList = Get-MgUser
$MGUserList.count
$MGUserList[5]
$MGUserList[5] | Get-Member
$MGUserList[5].Settings
$MGUserList[5].Mail
$MGUserList[5] | Select-Object *

# From Graph Explorer Snippets
Get-MgUser -ConsistencyLevel eventual

## PnP.PowerShell examples
# Retrieve all SharePoint sites
Invoke-PnPGraphMethod https://graph.microsoft.com/v1.0/sites?search=*

Invoke-PnPGraphMethod -Url "sites?search=*"
$sitelist = Invoke-PnPGraphMethod https://graph.microsoft.com/v1.0/sites?search=*

# Process the response
if ($sitelist) {
    foreach ($site in $sitelist.value) {
        $siteId = $site.id
        $siteName = $site.name
        $siteUrl = $site.webUrl

        # Create a custom object for the SharePoint site
        [PSCustomObject]@{
            TypeName = "TKMGSite"
            SiteName = $siteName
            SiteUrl = $siteUrl
            SiteId = $siteId
        }
    }
}

$sitelist.count
$sitelist.value
$sitelist.value.count
$sitelist.value[5]
$sitelist.value[5] | Get-Member
# Retrieve all users
$userlist = Invoke-PnPGraphMethod -Url "https://graph.microsoft.com/v1.0/users"
Invoke-PnPGraphMethod -Url "users"

# Number of Properties
($userlist.value[5] | Get-Member -MemberType NoteProperty).Count

# Process the response
$userlist.value
$userlist.value[5]

# Beta API
$userlist = Invoke-PnPGraphMethod -Url "https://graph.microsoft.com/beta/users"

# Number of Beta properties
($userlist.value[5] | Get-Member -MemberType NoteProperty).Count