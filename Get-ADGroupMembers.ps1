Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
<# 
Requirements: 
    * Must launch PowerShell or ISE with Run as Administrator
    * Set execution policy via command window first (minimum: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process)

Variables - You may hard code variables for ease of use. If left empty you will be prompted when script is executed.
    $LoginEmailAddress = 'paul@paulspoerry.com'
    $LoginEmailPassword = 'groovypwbro'
    $DistrosToSearch = '*developers*'   #may contain wildcards for multiple groups using asterick (*)
    $FileNameForCSV = 'email-recipients.csv'
    $ConnectionUri = 'https://outlook.office365.com/powershell-liveid/' #default for o365 / non-onprem
#>
$LoginEmailAddress = ''
$LoginEmailPassword = ''
$DistrosToSearch = ''
$FileNameForCSV = ''
$ConnectionUri = ''


# --- don't edit beyond here --- 
if ( !($PSVersionTable.PSVersion.Major -ge 3) ) { 
    Write-Error 'You must use PowerShell v3 or greater!'
    return
}


If ( (get-ExecutionPolicy) -ne "RemoteSigned" ) {
    Try {
        # If execution policy is anything other than Restricted then attemp to force RemoteSigned scoped for this process only
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
    } Catch {
        Write-Host "Current execution policy is $execPolicy.ToString(); script unable to run. Please execute the following and re-run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process" -foregroundcolor red
        return
    }
}


while([string]::IsNullOrWhiteSpace($LoginEmailAddress)){
    $LoginEmailAddress = Read-host 'Enter Username (or q to exit)'
    if($LoginEmailAddress -eq 'q'){exit}
}
if ([string]::IsNullOrWhiteSpace($LoginEmailPassword)) { 
    while([string]::IsNullOrWhiteSpace($LoginEmailPassword)){
        $LoginEmailPassword = Read-host 'Enter Password (or q to exit)' -AsSecureString
        $pw = $LoginEmailPassword
        if($LoginEmailPassword -eq 'q'){exit}
    }
} else {
    $pw = convertto-securestring -AsPlainText -Force -String $LoginEmailPassword
}
while([string]::IsNullOrWhiteSpace($DistrosToSearch)){
    $DistrosToSearch = Read-host 'Enter distribution group(s) to search [use * for wildcard] (or q to exit)'
    if($DistrosToSearch -eq 'q'){exit}
}
while([string]::IsNullOrWhiteSpace($FileNameForCSV)){
    $FileNameForCSV = Read-host 'Enter CSV filename for output (or q to exit)'
    if($FileNameForCSV -eq 'q'){exit}
}
while([string]::IsNullOrWhiteSpace($ConnectionUri)){
    $ConnectionUri = Read-host 'Enter connection URI (or q to exit)'
    if($ConnectionUri -eq 'q'){exit}
}


Write-Host ("     Creating credential object")
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $LoginEmailAddress, $pw

Write-Host ("     Attempting exchange connection")
Try
{
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $cred -Authentication Basic -AllowRedirection -ErrorAction Stop
    Write-Host ("     Connection successful")
} Catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Host ("ErrorMessage: " + $ErrorMessage + " | FailedItem: " + $FailedItem)
    return
}

Write-Host ("     Importing required Exchange cmdlet(s)")
Import-PSSession $session -CommandName Get-Dist* -AllowClobber -DisableNameChecking -verbose:$false | Out-Null 

$progress = 0
$distroGroups = Get-DistributionGroup -Identity $DistrosToSearch -RecipientTypeDetails 'MailUniversalDistributionGroup'
$groupMembers = foreach($group in $distroGroups) {
    $progress++
    #Write-Host "Scanning group $group"
    Write-Progress -activity “Iterating distribution groups” -status "Scanned: $progress of $($distroGroups.Count)" -PercentComplete (($progress / $distroGroups.count)*100) -Id 1 -CurrentOperation "Scanning group: $group"
    #Start-Sleep -Milliseconds 200 #debug timer
    Get-DistributionGroupMember $group.Identity -ResultSize "Unlimited" | Select DisplayName, PrimarySmtpAddress, Title, Department, Office, City, @{n='Phone';e={$_.Phone -replace '\+',' '}}, @{n='DistributionGroup';e={$group.Name}}, @{n='Group Managed By';e={$group.ManagedBy}}
}
$groupmembers | Sort Group,User | Export-Csv $FileNameForCSV -NoTypeInformation

#Close remote sessions
get-pssession | remove-pssession
Write-Host "     Closing Session`r`n`r`n"

Write-Host "Complete! Your results are in $FileNameForCSV."

