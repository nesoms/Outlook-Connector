#########################################################################################
## Outlook Commands by Example
#########################################################################################
 
 
#########################################################################################
## Connect to Outlook
 
# Outlook Connection
$Outlook = New-Object -ComObject Outlook.Application
 
## Listing Folders in Outlook (Shows Email, Calendar, Tasks etc)
$OutlookFolders = $Outlook.Session.Folders.Item(1).Folders
$OutlookFolders | ft FolderPath
 
## Using Default 
$OutlookDeletedITems = $Outlook.session.GetDefaultFolder(3)
$outlookOutbox = $Outlook.session.GetDefaultFolder(4)
$OutlookSentItems = $Outlook.session.GetDefaultFolder(5)
$OutlookInbox = $Outlook.session.GetDefaultFolder(6)
$OutlookCalendar = $Outlook.session.GetDefaultFolder(9)
$OutlookContacts = $Outlook.session.GetDefaultFolder(10)
$OutlookJournal = $Outlook.session.GetDefaultFolder(11)
$OutlookNotes = $Outlook.session.GetDefaultFolder(12)
$OutlookTasks = $Outlook.session.GetDefaultFolder(13)
 
 
#########################################################################################
## Inbox Folders
 
# List all Folders 
$Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Folders | ft FullFolderPath 
 
# Create folder
$Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Folders.Add("Scripts Received")
 
# Delete Folder
$OutlookFolderToDelete = $Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Folders.Item("Scripts Received")
$OutlookFolderToDelete.Delete()
 
 
#########################################################################################
## Inbox Email
 
## Navigating to Sub folder of Inbox called Daily Tasks
$Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Folders.Item("Daily Tasks")
 
# Read All Emails in a Folder Path Inbox -&amp;gt; SPAM Mail
$EmailsInFolder = $Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Folders.Item("SPAM Folder").Items
$EmailsInFolder | ft SentOn, Subject, SenderName, To, Sensitivity -AutoSize -Wrap
 
# Send an Email from Outlook
$Mail = $Outlook.CreateItem(0)
$Mail.To = "stephen@badseeds.local"
$Mail.Subject = "Action"
$Mail.Body ="Pay rise please"
$Mail.Send()           
 
# Delete an Email from the folder Inbox with Subject Title "Action"
$EmailInFolderToDelete = $Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Items
$EmailInFolderToDelete | ft SentOn, Subject, SenderName, To, Sensitivity -AutoSize -Wrap
$EmailToDelete = $EmailInFolderToDelete | Where-Object {$_.Subject -eq "Action"}
$EmailToDelete.Delete()
 
# Delete All Emails in Folder.Items
$EmailsInFolderToDelete = $Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Folders.Item("SPAM Folder").Items
foreach ($email in $EmailsInFolderToDelete)
    {
        $email.Delete()
    }
 
# Move Emails from Inbox to Test folder
$EmailIToMove = $Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Items
$EmailIToMove | ft SentOn, Subject, SenderName, To, Sensitivity -AutoSize -Wrap
$NewFolder = $Outlook.Session.Folders.Item(1).Folders.Item("Inbox").Folders.Item("test")
 
FOREACH($Email in $EmailIToMove )
    { 
        $Email.Move($NewFolder) 
    }
 
  
#########################################################################################
## Calender
 
## Connect to Calendar 
$OutlookCalendar = $Outlook.session.GetDefaultFolder(9)
 
# Read Calendar 
$OutlookCalendar.Items | ft subject, start
 
# Create New Calendar Item
$NewEvent = $Outlook.CreateItem(1)
$NewEvent.Subject = "Timmys Birthday"
$NewEvent.Start = [datetime]”10/9/2014";
$NewEvent.save()
 
# Create re-occuring Event
$NewEvent = $Outlook.CreateItem(1)
$NewEvent.Subject = "Timmys Birthday"
$NewEvent.Start = [datetime]”10/9/2014"
$Recur = $NewEvent.GetRecurrencePattern()
$Recur.Duration=1440
$Recur.Interval=12
$Recur.RecurrenceType=5
$Recur.Noenddate=$TRUE
$NewEvent.save()
 
# Delete Event - Timmys Birthday
$TimmyCalendar = $OutlookCalendar.Items | WHERE {$_.Subject -eq "Timmys Birthday"}
$TimmyCalendar.Delete()
 
 
#########################################################################################
## Tasks
 
# Read Tasks
$OutlookTasks = $Outlook.session.GetDefaultFolder(13).Items
$OutlookTasks | ft Subject, Body
 
# Create a task
$newTaskObject =  $Outlook.CreateItem("olTaskItem")
$newTaskObject.Subject = "New Task"
$newTaskObject.Body = "This is the main text"
$newTaskObject.Save()
 
# Delete a task
$OutlookTasks = $Outlook.session.GetDefaultFolder(13).Items
$DeleteTask = $OutlookTasks | Where-Object {$_.Subject -eq "New Task"}
$DeleteTask.Delete()
 
# Edit a task
$OutlookTasks = $Outlook.session.GetDefaultFolder(13).Items
$Task = $OutlookTasks | Where-Object {$_.Subject -eq "New Task"}
$Task.Body = "Updated Results"
$Task.Save()
 
 
#########################################################################################
## Contacts
 
# Read Contacts
$OutlookContacts = $Outlook.session.GetDefaultFolder(10).items
$OutlookContacts| Format-Table FullName,MobileTelephoneNumber,Email1Address
 
# Add a Contact
$OutlookContacts = $Outlook.session.GetDefaultFolder(10)
$NewContact = $OutlookContacts.Items.Add()
$NewContact | gm
$NewContact.FullName = "John"
$NewContact.Email1Address = "John@Badseeds.local"
$NewContact.Save()
 
# Delete Contact Full Name - "John"
$OutlookContacts = $Outlook.session.GetDefaultFolder(10).items
$DeleteJohn = $OutlookContacts | Where-Object {$_.FullName -eq "John"}
$DeleteJohn.Delete()
 
# Update Contact
$OutlookContacts = $Outlook.session.GetDefaultFolder(10).items
$John = $OutlookContacts | Where-Object {$_.FullName -eq "John"}
$John.CompanyName = "BadSeeds"
$John.Save()