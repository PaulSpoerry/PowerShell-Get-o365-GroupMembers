# PowerShell-Get-o365-GroupMembers
PowerShell script to retrieve user information from Office 365 distribution group(s) and export the data to a CSV file. The following information is returned:

 - DisplayName
 - PrimarySmtpAddress
 - Title
 - Department
 - Office
 - City
 - Phone
 - DistributionGroup
 - Group Managed By

#Requirements

 - Office 365
	 - Should work for an on-premise Exchange server but since I don't have one I'm unable to test that. 
 - PowerShell v3+
 - Must launch PowerShell or PowerShell ISE with 'Run as Administrator'
 - Set appropriate Windows PowerShell Execution Policy for your environment via command line 
	 - Minimum: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

#WTF is this for?
In short - using PowerShell to query o365 distribution group(s) in order to obtain user information. In my case I needed it for SharePoint. You can read more [on the Wiki](https://github.com/PaulSpoerry/PowerShell-Get-o365-GroupMembers/wiki#wtf-is-this-for). 
#Usage
You may edit the script and hard code variables for ease of use. If left empty you will be prompted when script is executed. If you want to hard code them for ease of use they are defined in the script as follows:

    $LoginEmailAddress = 'paul@paulspoerry.com'
    $LoginEmailPassword = 'groovypwbro'
    $DistrosToSearch = '*developers*' #may contain wildcards for multiple groups using asterick (*)
    $FileNameForCSV = 'email-recipients.csv'
    $ConnectionUri = 'https://outlook.office365.com/powershell-liveid/' #default for o365

If not already modified set the execution policy (the following is scoped to only modify it for the current process) as follows:

    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

Then execute the script to retrieve the results:

    .\Get-ADGroupMembers.ps1

Your results will be output to the CSV file you defined in the script or enter during execution. The file will be output to the location you executed the script from in your PowerShell session. 
#ToDo's
It'll get the job done but could certainly be cleaned up/optimized and some better error handling could be implemented. For instance, if the script runs successfully it will clean up any remote connections. If it fails the connection will not be automatically closed by the script (sry/not sry/lazy); the server WILL automagically timeout and close the session after 15 minutes though.  Feel free to contribute as you see fit but this should be a decent start. 
