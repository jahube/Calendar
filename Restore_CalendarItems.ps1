
Install-Module -Name ExchangeOnlineManagement

Connect-ExchangeOnline

   $USER = "Modify user"     #  <-- Modify

# OutGridView Calendar Items - select + Restore

Get-RecoverableItems -Identity $USER -FilterItemType "IPM.Appointment" | OGV -PassThru | Restore-RecoverableItems


$Logpath = "C:\temp"

# DELETED_Calendar_Folders

$FolderStats = Get-MailboxFolderStatistics $user -IncludeSoftDeletedRecipients -IncludeAnalysis

$DeletedItems_folder = $FolderStats.where( {$_.FolderType -eq "DeletedItems"})

$Deleted_Folders = $FolderStats.where( {$_.ParentFolderId -eq $DeletedItems_folder.FolderId })

$FolderStats | Fl * > "$Logpath\DELETED_Calendar_Folders.txt"


# ALL_Calendar_Folders

$FolderStats = (Get-MailboxFolderStatistics $user).where( {$_.ContainerClass -eq "IPF.Appointment"})

$FolderStats | Fl * > "$Logpath\ALL_Calendar_Folders.txt"


<#

$ContainerClass_Array = @("IPM.Appointment","IPF.Appointment")

$recoverable_Items = Get-RecoverableItems -Identity $USER -FilterItemType "IPM.Appointment"

$recoverable_Items.where( {$_.ItemClass -in $ContainerClass_Array}) | OGV -PassThru | Restore-RecoverableItems

$recoverable_Items.where( {$_.ItemClass -in $ContainerClass_Array}) | fl *

#>