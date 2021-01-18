# MicrosoftTeamsChatExportTool  
This tool is for exporting Microsoft teams chat data from any user of the organization, with a pst file.
The admin who will be connecting to powershell must have the eDiscovery Manager and Compliance administrator permissions assigned.

For users hosted onprem keep in mind that if you have never performed a search for a onpremises mailbox you will need to read this article
https://docs.microsoft.com/en-us/microsoft-365/compliance/search-cloud-based-mailboxes-for-on-premises-users?view=o365-worldwide

You will need to engage Microsoft support in order whitelist your tenant for a search over a onprem user mailbox.

# How to run  

Download the  PS1 file and open powershell as an admin
``` powershell
.\TeamsExportChatTool.ps1
```
