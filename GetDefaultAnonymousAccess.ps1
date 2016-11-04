#This Script sets the 'Default' and 'Anonymous' Access to NONE for any mail box where $folder.AccessRights does not currently equal 'None'
#https://technet.microsoft.com/en-us/library/ff522363%28v=exchg.160%29.aspx

#Log File
$logfile = 'C:\Temp\ExchangeFolders\Logs\ExchangeAccess.log'
$Title = "No,Mailbox,FolderPath,Default,DefaultAccess,Anonymous,DefaultAccess`n"
$Title | New-Item $logfile -type file -force
[int64]$count = 0

#Import Mailboxes to Search
#Get all User Accounts which have the attribute Resources, this is all PBEAs

$ADUsers = get-aduser -f{samaccountname -like '*'} -Properties mail

#Go Though each Mailbox
foreach($Mailbox in $ADUsers)
{
    #Append Count
    $count += 1
    #Get List of all Folders in Mailbox
    $MailboxFolders =  Get-MailboxFolderStatistics $mailbox.mail
    #Go Through each Folder
    foreach($folder in $MailboxFolders)
    {
        #Format Folder Path, swap / to \, e.g Travis.DRAKE@police.vic.gov.au:\Inbox\Archived
        $folderpath = $($Mailbox.mail + ":" + $folder.FolderPath -replace('/','\'))
        #Get the Permissions listed on the Folder
        $folderPermissions = Get-MailboxFolderPermission -Identity $folderpath -ErrorAction SilentlyContinue
        #Pull the Default Access
        $Default = $folderPermissions | where {$_.user -eq 'Default'}
        #Pull the Anonymous Access
        $Anonymous = $folderPermissions | where {$_.user -eq 'Anonymous'}
        #Log to Console
        Write-Host "$(5596-$count) - Mailbox $($Mailbox.Name) - $folderpath has User $($Default.User) with access: $($Default.AccessRights) and User $($Anonymous.User) with access: $($Anonymous.AccessRights)"
        #Log to CSV
        $log = "'$count','$($Mailbox.Name)','$folderpath','$($Default.User)','$($Default.AccessRights)','$($Anonymous.User)','$($Anonymous.AccessRights)'"
        $log | Add-Content $logfile
        #Add a Sleep to give the exchange server a little break
        sleep 0.1
    }
}