
$Outlook = New-Object -ComObject Outlook.Application
 
 $Outlook.Session.Folders.Item(1).Folders.Item("Calendar").Folders | ft FullFolderPath  


 $Outlook.Session.Folders.Item(1).Folders.Item("Calendar").Folders.Item("Freinds Reach Out")
