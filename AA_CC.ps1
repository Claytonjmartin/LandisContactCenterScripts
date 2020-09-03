#This script will configure a Landis Contact Center queue resource account to be used as a transfer option in a MS Teams Auto Attendant.
#The admin running this script will be promted to make selections based on the desired call flow of the Auto Attendant.

#Prerequisites:
#Active connection to Skype for Business online powershell
#Teams Auto Attendant created
#Resource account created for Landis Contact Center Queue

#Author: Clayton Martin, Landis Technologies LLC
#This script is provided "As Is" without any warranty of any kind. In no event shall the author be liable for any damages arising from the use of this script.

$AAS = Get-CsAutoAttendant
$RAS = Get-CsOnlineApplicationInstance
$Objects = @("After Hours Call Flow", "Business Hours Call Flow", "Operator", "Holiday")
$Actions = @("Redirect Call", "Menu Options")
$dial_Key = @(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)

#Select AA
$global:selection = $null
Clear-Host
if (!$aas.count) {
    $AA_Identity = $AAS.Identity
    $AAName  = $AAS.Name
    Read-Host "Editing $AAName. Press ENTER to continue..."
}
if ($AAS.count) {
    Do {
        Write-Host 'Auto Attendant to Edit:'

        for ($i = 0; $i -lt $AAS.count; $i++) {
            Write-Host -ForegroundColor Cyan "  $($i+1)." $AAS.Name[$i]
        }
        Write-Host # empty line
        $global:ans = (Read-Host 'Please enter selection') -as [int]

    } While ((-not $ans) -or (0 -gt $ans) -or ($AAS.Count -lt $ans))

    $global:selection = $AAS[$ans - 1]
    $AA_Identity = $global:selection.Identity
}
#Select Resource Account
$global:selection = $null
Clear-Host
Do {
    Write-Host 'Resource Account to transfer to:'

    for ($i = 0; $i -lt $RAS.count; $i++) {
        Write-Host -ForegroundColor Cyan "  $($i+1)." $RAS.DisplayName[$i]
    }
    Write-Host # empty line
    $global:ans = (Read-Host 'Please enter selection') -as [int]

} While ((-not $ans) -or (0 -gt $ans) -or ($RAS.Count -lt $ans))

$global:selection = $RAS[$ans - 1]
$RA_ObjectID = $global:selection.ObjectId

#Select Object
$global:selection = $null
Clear-Host
Do {
    Write-Host 'Auto Attendant Object to Edit:'

    for ($i = 0; $i -lt $Objects.count; $i++) {
        Write-Host -ForegroundColor Cyan "  $($i+1)." $Objects[$i]
    }
    Write-Host # empty line
    $global:ans = (Read-Host 'Please enter selection') -as [int]

} While ((-not $ans) -or (0 -gt $ans) -or ($Objects.Count -lt $ans))

$global:selection = $Objects[$ans - 1]
$Object = $global:selection

#Select Actions
Function Get-AAAction {
    $global:selection = $null 
    Clear-Host
    Do {
        Write-Host 'Auto Attendant Action:'

        for ($i = 0; $i -lt $Actions.count; $i++) {
            Write-Host -ForegroundColor Cyan "  $($i+1)." $Actions[$i]
        }
        Write-Host # empty line
        $global:ans = (Read-Host 'Please enter selection') -as [int]

    } While ((-not $ans) -or (0 -gt $ans) -or ($Actions.Count -lt $ans))

    $global:selection = $Actions[$ans - 1]
    return $global:selection
}

#Select Key
Function Get-AAKey {
    $global:selection = $null
    Clear-Host
    Do {
        Write-Host 'Auto Attendant Dial Key to transfer to Resource Account:'

        for ($i = 0; $i -lt $dial_Key.count; $i++) {
            Write-Host -ForegroundColor Cyan "  $($i+1)." $dial_Key[$i]
        }
        Write-Host # empty line
        $global:ans = (Read-Host 'Please enter selection') -as [int]

    } While ((-not $ans) -or (0 -gt $ans) -or ($dial_Key.Count -lt $ans))

    $global:selection = $dial_Key[$ans - 1]
    return "Tone" + $global:selection
}


$aa = Get-CsAutoAttendant -Identity $AA_Identity

If ($Object -like "Holiday") {
    $HCF = $AA.CallFlows | where { $_.Menu.Name -like "Holiday call flow" }
    #Select Holiday
    $global:selection = $null
    Clear-Host
    Do {
        Write-Host 'Holiday to Edit:'
        for ($i = 0; $i -lt $HCF.count; $i++) {
            Write-Host -ForegroundColor Cyan "  $($i+1)." $HCF.Name[$i]
        }
        Write-Host # empty line
        $global:ans = (Read-Host 'Please enter selection') -as [int]
    } While ((-not $ans) -or (0 -gt $ans) -or ($HCF.Count -lt $ans))
    $global:selection = $HCF[$ans - 1]
    $HN = $global:selection.Name
    $holiday = $HCF | where Name -Like $HN | select -First 1
    $Callable_Entity = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
    $holidayMO = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -CallTarget $Callable_Entity -DtmfResponse Automatic
    $holiday.Menu.MenuOptions = $holidayMO
}

If ($Object -like "Operator") {
    $aa.Operator = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
}

If ($Object -like "After Hours Call Flow") {
    $AHCF = $AA.CallFlows.menu | where Name -like "After hours call flow" | select -First 1
    $AHCFM = $AHCF.MenuOptions | where DtmfResponse -Like "Automatic" | select -First 1
    If ($AHCFM -ne $null) {
        $action = Get-AAAction
        if ($action -like "Redirect Call") {
            $AHCFM.Action = "TransferCallToTarget"
            $AHCFM.CallTarget = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
        }
        if ($action -like "Menu Options") {
            $key = Get-AAKey
            $AHCFM.Action = "TransferCallToTarget"
            $AHCFM.DtmfResponse = $key
            $AHCFM.CallTarget = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
            $AHCF.Prompts = New-CsAutoAttendantPrompt -TextToSpeechPrompt "This is a default greeting. Please edit in the Teams Admin Center."                
        }
    }
    If ($AHCFM -eq $null) {
        $key = Get-AAKey
        $mo = $AHCF.MenuOptions | where DtmfResponse -Like $key | select -First 1
        if ($mo -ne $null) {
            $mo.Action = "TransferCallToTarget"
            $mo.CallTarget.Type = "ApplicationEndpoint"
            $mo.CallTarget.ID = $RA_ObjectID
        }

        if ($mo -eq $null) {
            $Callable_Entity = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
            $menuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse $key -CallTarget $Callable_Entity
            if ($AHCF.MenuOptions -ne $null) {
                $AHCF.MenuOptions.Add($menuOption)
            }
            if ($AHCF.MenuOptions -eq $null) {
                $AHCF.MenuOptions = $menuOption
            }
        }  
    }
}

If ($Object -like "Business Hours Call Flow") {
    $AHCF = $AA.DefaultCallFlow.menu | where {$_.Name -like "MSPhone_AutoAttendant_*" -or $_.Name -like "Business hours call flow"} | select -First 1
    $AHCFM = $AHCF.MenuOptions | where DtmfResponse -Like "Automatic" | select -First 1
    If ($AHCFM -ne $null) {
        $action = Get-AAAction
        if ($action -like "Redirect Call") {
            $AHCFM.Action = "TransferCallToTarget"
            $AHCFM.CallTarget = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
        }
        if ($action -like "Menu Options") {
            $key = Get-AAKey
            $AHCFM.Action = "TransferCallToTarget"
            $AHCFM.DtmfResponse = $key
            $AHCFM.CallTarget = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
            $AHCF.Prompts = $AHCF.Prompts = New-CsAutoAttendantPrompt -TextToSpeechPrompt "This is a default greeting. Please edit in the Teams Admin Center."
        }
    }
    If ($AHCFM -eq $null) {
        $key = Get-AAKey
        $mo = $AHCF.MenuOptions | where DtmfResponse -Like $key | select -First 1
        if ($mo -ne $null) {
            $mo.Action = "TransferCallToTarget"
            $mo.CallTarget.Type = "ApplicationEndpoint"
            $mo.CallTarget.ID = $RA_ObjectID
        }

        if ($mo -eq $null) {
            $Callable_Entity = New-CsAutoAttendantCallableEntity -Identity $RA_ObjectID -Type ApplicationEndpoint
            $menuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse $key -CallTarget $Callable_Entity
            if ($AHCF.MenuOptions -ne $null) {
                $AHCF.MenuOptions.Add($menuOption)
            }
            if ($AHCF.MenuOptions -eq $null) {
                $AHCF.MenuOptions = $menuOption
            }
        }        
    }
}

Try{
    Set-CsAutoAttendant -Instance $aa
} catch {Throw "Cannot Edit Auto Attendant. " + $error[0]}
Write-Host "Auto attendant configuration changed sucessfully!"