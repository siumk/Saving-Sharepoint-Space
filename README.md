# Sharepoint-No-Versioning

Enable NO Versioning on Sharepoint and remove old versioned files

This project is aimed to provide a simple script to save $$ by enabling the hidden option of ZERO versioning on Microsoft Sharepoint site and removing ALL old versioned files.

The Powershell script can recursively traverse across a huge number of folders, subfolders and files.

The Powershell script can be run from a Windows Client with Sharepoint user credential and without the needs of Sharepoint admin console.

Features :
* Can be run from a Windows Client with Sharepoint site user credential
* Recursively traverse across folders, subfolders and files
* Can work on large sized site
* Supports O365 Sharepoint
* Supports Teams files share on Sharepoint
* Save space Save money

HOW-TO :

1. Firstly, understand the issue on how your $pace is slipping through your fingers

   https://techcommunity.microsoft.com/t5/microsoft-onedrive-blog/new-updates-to-onedrive-and-sharepoint-team-site-versioning/ba-p/204390

2. Secondly, install the "SharePoint Online Client Components SDK" 

   https://www.microsoft.com/en-us/download/details.aspx?id=42038

   The installation will get all the required dlls at 16 hive location.

   Here is the location where all the needed dlls.

   C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\

3. Then, change the Site URL in the script

   $SiteURL="*** SET YOUR SITE NAME ***"

   Sample site names:
   $SiteURL="https://XXX.sharepoint.com/teams/ABC" # Traditional Sharepoint
   $SiteURL="https://XXX.sharepoint.com/sites/msteams_XXXXXX" # Teams Sharepoint

4. Run the script with PowerShell

5. Enter Sharepoint site user login name XXX@YYY.ZZZ and password

6. Sit back and Watch it run

7. Finally, verify the versioning setting of the Sharepoint Site

   https://support.microsoft.com/en-us/office/enable-and-configure-versioning-for-a-list-or-library-1555d642-23ee-446a-990a-bcab618c7a37


[Optional] 

Read the following very helpful article on Using CSOM in PowerShell scripts with Office 365 by Chris O'Brien
http://www.sharepointnutsandbolts.com/2013/12/Using-CSOM-in-PowerShell-scripts-with-Office365.html


References :

https://www.sharepointdiary.com/2018/08/sharepoint-online-powershell-to-disable-versioning.html
https://www.sharepointdiary.com/2016/02/sharepoint-online-delete-version-history-using-powershell.html
http://www.sharepointnutsandbolts.com/2013/12/Using-CSOM-in-PowerShell-scripts-with-Office365.html



** DISCLAIMER **
The author will not accept any responsibility for loss of computer data or software, however caused, including any alleged loss sustained due to the execution of the script. The author is not liable for any direct, indirect, consequential, incidental or punitive damages arising as a result of your use of or inability to download the script and / or instruction thereof and / or use of or inability to execute the script. The author provides no warranties whatsoever, express or implied that the instruction of any and all of our scripts is error-free and / or up-to-date and that script will be 100 % error-free.
