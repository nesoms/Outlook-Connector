

Add-Type -assembly "Microsoft.Office.Interop.Outlook"

function add-appointment(){
<#
.SYNOPSIS
Creates and appointment in outlook
.DESCRIPTION
Creates an appointment in outlook in the default calendar folder
.PARAMETER TimeDate
Time and Date of the appointment. Multiple formats are accepted. Please use the one common to your region. It will create the appointment for today if no date is specified. 
.PARAMETER Location 
The message to be send.
.PARAMETER ReminderMinutesBeforeStart
Sets when to send reminder for appointment. Default is 15 minutes. 1440 minutes are in a day and 10080 minutes in a week. Set it to 0 for no reminders.
.PARAMETER Subject
The subject of the appointment.
.EXAMPLE
add-appointment -TimeDate "1/4/2015 4:30 pm" -Subject "Little Mermaid" -location "Cameo Theater" -ReminderMinutesBeforeStart 0
This example adds a appointment on 1/4/2015 at 4:30 PM with subject Little mermaid and no reminder.  
.NOTES
Requires Windows PowerShell v2 or later and Microsoft Office Outlook 2010 or later.
Outlook cannot be running when running this script.Some versions of Outlook will not close the session. Please go to taskbar and close manually in that case.  
#>



[CmdletBinding()]
param (
[parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
$TimeDate, 
[parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
$Subject ,
$location,
$body,
[int]$ReminderMinutesBeforeStart = "15"
)


$AppointmentTime = get-date $Timedate
#Check whether Outlook is installed
if (-not (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE")) {
Write-Warning "Outlook is not installed. You need to install Outlook 2010 or later to use this Script."
break
}
 



$outlook = new-object -com Outlook.Application
$calendar = $outlook.Session.GetDefaultFolder(9) # == olFolderCalendar 
$calendar =  $Outlook.Session.Folders.Item(1).Folders.Item("Calendar").Folders.Item("Freinds Reach Out")
$appt = $calendar.Items.Add(1) # == olAppointmentItem 
$appt.Start = [datetime]$AppointmentTime
$appt.Subject = $subject
$appt.Location = $location
$appt.ReminderSet = $true
$appt.Body = $Body
$appt.ReminderMinutesBeforeStart =$ReminderMinutesBeforeStart
$appt.GetRecurrencePattern


#################


$Recur = $appt.GetRecurrencePattern()
$Recur.Duration=30

$Recur.RecurrenceType=2
$Recur.Interval=3
$Recur.Noenddate=$TRUE
#$Recur.Regenerate=$TRUE




##################






$appt.Save() 
$outlook.Quit() 



 





}


#add-appointment -TimeDate "11/12/2019 4:30 pm" -Subject "Reach-out: Steven Judd" -location "Cameo Theater" -ReminderMinutesBeforeStart 0


$Contact_List = Import-Csv "C:\Users\scotn\OneDrive\Documents\Scripts\Outlook Connector\APPT-Contacts.csv"
#$Contact_List = Import-Csv "C:\Users\scotn\OneDrive\Documents\Scripts\Outlook Connector\Test-APPT-Contacts.csv"


foreach( $Item in $Contact_List){

$FirstName = $Item.'First Name'
$LastName = $Item.'Last Name'
$Phone = $Item.'Mobile Phone'
$Time = "11/12/2019 8:30 am"
$Subject = "Reach-Out: " + $FirstName + " " + $LastName
$Location = "SMS Message"
$Body = "Cell Phone:   " + $Phone.ToString()

########################## Generate Date

$StartDate = Get-Date -Date 2019-1-01
$EndDate = Get-Date -Date 2019-06-01

$RangeInDays = 0..(($EndDate - $StartDate).Days)

$DaysToAdd = Get-Random -InputObject $RangeInDays

$RandDate = $StartDate.AddDays($DaysToAdd)


########################################



$time = ($RandDate).ToString("MM/dd/yyyy") + " 8:30 am"


$FirstName 
$LastName
$phone
$time 



#$Body = "<a href='sms:" + $Phone.ToString() +  "'>" +  $Phone.ToString() +  "</a>"
add-appointment -TimeDate $Time -Subject $Subject -location $Location -Body $Body -ReminderMinutesBeforeStart 0

}


