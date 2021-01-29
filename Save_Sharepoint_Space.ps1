#
# Firstly, install the SharePoint Online Client Components SDK. This installation will get all the required dlls at 16 hive location.
# Here is the location where all the needed dlls.
#
# C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\
#
# [Optional] Read the following very helpful article on Using CSOM in PowerShell scripts with Office 365 by Chris O'Brien
# http://www.sharepointnutsandbolts.com/2013/12/Using-CSOM-in-PowerShell-scripts-with-Office365.html
#
# Then, change the Site URL below - Please see the sample site name
#
    
# Config Parameters
#$SiteURL="https://XXX.sharepoint.com/teams/ABC" # Traditional Sharepoint
#$SiteURL="https://XXX.sharepoint.com/sites/msteams_XXXXXX" # Teams Sharepoint
$SiteURL="*** SET YOUR SITE NAME ***"


# Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# Setup root folder
$ListName="Documents"
$pos = $SiteURL.IndexOf(".com/")
$RootFolder = $SiteURL.Substring($pos+4)

#Setup Credentials to connect
# $UserName="XXXXX"
# $Password ="XXXXX"
# $Cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,(ConvertTo-SecureString $Password -AsPlainText -Force))
$Cred = Get-Credential
$Cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName,$Cred.Password)
#$Global:Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username,$Cred.Password)

#Function to Disable Versioning on All Document Libraries in a SharePoint Online Site
Function Disable-SPOVersionHistory()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL
    )
    Try {
        Write-host -f Yellow "Processing site:"$SiteURL
 
        #Get the site, subsites and lists from given site
        $Web = $Ctx.web
        $Ctx.Load($Web)
        $Ctx.Load($Web.Lists)
        $Ctx.Load($web.Webs)
        $Ctx.executeQuery()
 
        #Array to exclude system libraries
        $SystemLibraries = @("Form Templates", "Pages", "Preservation Hold Library","Site Assets", "Site Pages", "Images",
                            "Site Collection Documents", "Site Collection Images", "Style Library")
         
        #Get All document libraries
        $DocLibraries = $Web.Lists | Where {$_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $SystemLibraries}
        ForEach($Library in $DocLibraries)
        {
            #disable versioning in each document library
            $Library.EnableVersioning = $False
            $Library.Update()
            $Ctx.ExecuteQuery()
            Write-host -f Green "`tVersioning has been turned OFF at '$($Library.Title)'"
        }
  
        #Iterate through each subsite
        ForEach ($Subweb in $Web.Webs)
        {
            #Call the function recursively
            Disable-SPOVersionHistory($Subweb.url)
        }
    }
    Catch {
        write-host -f Red "Error:" $_.Exception.Message
    }
# Reference 1: https://www.sharepointdiary.com/2018/08/sharepoint-online-powershell-to-disable-versioning.html
}

function CheckFiles([string]$targetFolder)
{
    $Query.FolderServerRelativeUrl = $targetFolder;

    $Query.ViewXml = "<View>" +
		     "<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>" +
		     "<Query>" +
			"<Where>" +
		     	"<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq>" +
			"<Neq><FieldRef Name='ContentType' /><Value Type='Text'>Folder</Value></Neq>" +
		     	"</Where>" +
		     "</Query>" +
#		     "<RowLimit>5</RowLimit>" +
		     "</View>";

    $ListItems = $List.GetItems($Query)
    $Ctx.Load($ListItems)
    $Ctx.ExecuteQuery() 

    write-host "Total Number of Files:"$ListItems.Count    
    # Loop through each file in the library
    Foreach($Item in $ListItems)
    {      
        # Get all versions of the file
        $Ctx.Load($Item.File)       
        $Ctx.Load($Item.File.Versions)
        $Ctx.ExecuteQuery()
    	write-host "File=" $Item.File.Name -f yellow "#=" $Item.File.Versions.count

        # Delete all versions of the found file
        If($Item.File.Versions.count -gt 0)
        {
            $Item.File.Versions.DeleteAll()
            $Ctx.ExecuteQuery()
            Write-host -f Green "Cleaned"
        }
     }

#Reference 2: https://www.sharepointdiary.com/2016/02/sharepoint-online-delete-version-history-using-powershell.html
}

function CheckFolder([string]$targetFolder)
{

    $Query.FolderServerRelativeUrl = $targetFolder;

    write-host "Check Folder:"$Query.FolderServerRelativeUrl

    $Query.ViewXml = "<View>" +
		     "<QueryOptions><ViewAttributes Scope='RecursiveAll'/></QueryOptions>" +
		     "<Query>" +
			"<Where>" +
		     	"<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>" +
			"<Eq><FieldRef Name='ContentType' /><Value Type='Text'>Folder</Value></Eq>" +
		     	"</Where>" +
		     "</Query>" +
#		     "<RowLimit>5</RowLimit>" +
		     "</View>";

    $ListItems = $List.GetItems($Query)
    $Ctx.Load($ListItems)
    $Ctx.ExecuteQuery() 

    write-host "Total Number of Subfolders:"$ListItems.Count    
    # Loop through each item
    $ListItems | ForEach-Object {
    	write-host "FR=" $_["FileRef"] "FLR=" $_["FileLeafRef"]
	CheckFiles($_["FileRef"])
	CheckFolder($_["FileRef"])
    }  
}

# main
Try {

    # Setup the context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Ctx.Credentials = $Cred

    # Call the function to disable versions on all document libraries
    Disable-SPOVersionHistory -SiteURL $SiteURL
  
    # Get the web and Library
    $Web=$Ctx.Web
    $List=$web.Lists.GetByTitle($ListName)

    $ListItems = $List.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()) 

    $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
   
    # CheckFolder("/teams/XXX_YYY/Shared Documents"); # Root folder reference
    CheckFolder($RootFolder);
}
Catch {
    write-host -f Red "Error Deleting version History!" $_.Exception.Message
}
# Source: https://sourceforge.net/p/saving-sharepoint-space/wiki/
