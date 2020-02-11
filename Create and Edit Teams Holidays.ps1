#This script will create or edit Microsoft Teams holiday schedules

#Prerequisites:
    #Connect PowerShell to Skype for Business online

#Author: Clayton Martin, Landis Technologies LLC
#This script is provided "As Is" without any warranty of any kind. In no event shall the author be liable for any damages arising from the use of this script.


#Parameter Options: https://kayaposoft.com/enrico/json/

$HolidayYear = "2021"
$Country = "usa"
$region = "pa"
$holidayType = "public_holiday"
#Time must be in 15min increments
$StartTime = "00:00"
$EndTime = "23:45"



#Get Existing Holiday Schedules
$ExistingHolidays = Get-CsOnlineSchedule

#Get New Holidays
$uri = "https://kayaposoft.com/enrico/json/v2.0?action=getHolidaysForYear&year=" + $HolidayYear + "&country=" + $Country + "&region=" + $region + "&holidayType=" + $holidayType
$Holidays = (Invoke-WebRequest -Method Get -Uri $uri).Content | ConvertFrom-Json
[System.Collections.ArrayList]$formattedHolidays = @()
foreach ($Holiday in $Holidays){
    $name = $Holiday.name.text -replace ",", ""
    $date = $holiday.date.day.ToString() + "/" + $holiday.date.month.ToString() + "/" + $holiday.date.year.ToString()
    $startdate = $date + " " + $StartTime
    $enddate = $date + " " + $EndTime
    $PCO = [PSCustomObject]@{
        Name = $name
        StartDateTime = $startdate
        EndDateTime = $enddate
    }
    $formattedHolidays.Add($pco) | Out-Null
}

#Check If Holiday Exists
[System.Collections.ArrayList]$HolidaysToEdit = @()
foreach ($formattedHoliday in $formattedHolidays){
    foreach ($ExistingHoliday in $ExistingHolidays){
        if ($ExistingHoliday.Name -eq $formattedHoliday.Name){
            $HolidaysToEdit.Add($formattedHoliday) | Out-Null
        }
    }
}
#Remove Existing Holidays From Being Created
for ($i = 0; $i -lt $HolidaysToEdit.Count; $i++){
    $formattedHolidays.Remove($HolidaysToEdit[$i])
}

#Create New Holidays
foreach ($formattedHoliday in $formattedHolidays){
    $name = $formattedHoliday.Name
    Write-Host "Would you like to add the $name Holiday?"
    $ReadHost = Read-Host "( y / n )"
    if ($ReadHost -eq "y"){
        $dtr = New-CsOnlineDateTimeRange -Start $formattedHoliday.StartDateTime -End $formattedHoliday.EndDateTime
        New-CsOnlineSchedule -Name $formattedHoliday.Name -FixedSchedule -DateTimeRanges @($dtr)
    }
}

#Edit Existing Holidays
foreach ($HolidayToEdit in $HolidaysToEdit){
    $name = $HolidayToEdit.Name
    Write-Host "Holiday $name already exists. Would you like to add a new Date and Time Schedule?"
    $ReadHost = Read-Host "( y / n )"
    if ($ReadHost -eq "y"){
        $DateTime = New-CsOnlineDateTimeRange -Start $HolidayToEdit.StartDateTime -End $HolidayToEdit.EndDateTime
        $schedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayToEdit.Name}
        if ($schedule.FixedSchedule.DateTimeRanges.Count -le 10){
            $schedule.FixedSchedule.DateTimeRanges += $DateTime
            Try{
                Set-CsOnlineSchedule -Instance $schedule -ErrorAction stop
            }
            Catch{
                $errorMessage = $_.exception.message
                Write-Host "Holiday $name`: $errorMessage Skipping....." -ForegroundColor Red

            }
        }
        if ($schedule.FixedSchedule.DateTimeRanges.Count -gt 10){
            Write-Host "Cannot add date and time range because there is a maximum of 10 already specified for $name. Please delete some schedules for this holiday and run this script again."
        }
    }
}
