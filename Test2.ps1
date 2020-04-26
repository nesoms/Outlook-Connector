$newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
$newApt.Start = [DATETIME]::Now
$newApt.End = [DATETIME]::Now.AddMinutes(60)
$newApt.Subject = "Test is a test"
$DayOfTheWeek = New-Object Microsoft.Exchange.WebServices.Data.DayOfTheWeek[] 1
$DayOfTheWeek[0] = [Microsoft.Exchange.WebServices.Data.DayOfTheWeek]::Saturday 
$newApt.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+WeeklyPattern([DATETIME]::Now, 1, $DayOfTheWeek);
$newApt.Recurrence.StartDate = [DATETIME]::Now
$newApt.Recurrence.NumberOfOccurrences = 10
$newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)