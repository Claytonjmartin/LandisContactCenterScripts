#This script through a series of menu prompts will failover a Landis Contact Center resource account to a Teams auto attendant/call queue or failback the resource account to Landis Contact Center.

#Prerequisites:
#Teams PowerShell Module installed
#Failover Teams Auto Attendant and/or call queue created
#Resource account created for Landis Contact Center Queue/IVR

#Author: Clayton Martin, Landis Technologies LLC
#This script is provided "As Is" without any warranty of any kind. In no event shall the author be liable for any damages arising from the use of this script.
$Module = Get-InstalledModule microsoftteams
if (!$Module) {
    Write-Host "Please install Teams PowerShell Module and try again."
}
if ($Module) {
    try {
        #Connect-MicrosoftTeams
        Clear-Host
        Write-Host "Getting Teams environment data......."
        $AAS = Get-CsAutoAttendant -WarningAction silentlyContinue
        $CQS = Get-CsCallQueue -WarningAction silentlyContinue
        $CQAID = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
        $AAAID = "ce933385-9390-45d1-9512-c8d228074e07"
        $LCCAID = "341e195c-b261-4b05-8ba5-dd4a89b1f3e7"
        $Actions = @("Failover Landis Contact Center Queue or IVR to Teams Call Queue or Auto Attendant", "Failback Teams Queue or Auto Attendant to Landis Contact Center", "Exit")
        $Options = @("Teams Auto Attendant", "Teams Call Queue")
    }
    catch { Throw $error[0] }
    #Select Action
    function show-menu {
        try {
            Write-Host "Getting Resource accounts. Please wait......"
            $appInstacnesWithAppId = @()
            $appInstances = Get-CsOnlineApplicationInstance
            foreach ($appInstance in $appInstances) {
                $details = Get-CsOnlineApplicationInstance -Identity $appInstance.ObjectId
                $appInstacnesWithAppId += $details 
            }
            $LCCRAS = $appInstacnesWithAppId | Where-Object { $_.ApplicationId -eq "341e195c-b261-4b05-8ba5-dd4a89b1f3e7" }
            $RAS = $appInstacnesWithAppId | Where-Object { ($_.ApplicationId -eq "11cd3e2e-fccb-42ad-ad00-878b93575e07") -or ($_.ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07") }
            $global:selection = $null 
            Clear-Host
            Do {
                Write-Host 'Main Menu:'

                for ($i = 0; $i -lt $Actions.count; $i++) {
                    Write-Host -ForegroundColor Cyan "  $($i+1)." $Actions[$i]
                }
                Write-Host # empty line
                $global:ans = (Read-Host 'Please enter selection') -as [int]

            } While ((-not $ans) -or (0 -gt $ans) -or ($Actions.Count -lt $ans))

            $global:selection = $Actions[$ans - 1]
        }
        catch { Throw $error[0] }
        if ($global:selection -eq "Exit") {
            break
        }
        try {
            if ($global:selection -eq "Failover Landis Contact Center Queue or IVR to Teams Call Queue or Auto Attendant") {
                #Select LCC Resource Account
                $global:selection = $null
                Clear-Host
                if ($LCCRAS.count -le 1) {
                    $LCCRAS_ObjectID = $LCCRAS.ObjectId
                    $RAName = $LCCRAS.DisplayName
                    Read-Host "Failing over Landis Contact Center Resource Account $RAName. Press ENTER to continue..."
                }
                if ($LCCRAS.count -ge 2) {
                    Do {
                        Write-Host 'Landis Contact Center Resource Account to Failover:'

                        for ($i = 0; $i -lt $LCCRAS.count; $i++) {
                            Write-Host -ForegroundColor Cyan "  $($i+1)." $LCCRAS.DisplayName[$i]
                        }
                        Write-Host # empty line
                        $global:ans = (Read-Host 'Please enter selection') -as [int]

                    } While ((-not $ans) -or (0 -gt $ans) -or ($LCCRAS.Count -lt $ans))

                    $global:selection = $LCCRAS[$ans - 1]
                    $LCCRAS_ObjectID = $global:selection.ObjectId
                }

                #Select failover option AA or CQ

                $global:selection = $null 
                Clear-Host
                Do {
                    Write-Host 'Failover to:'

                    for ($i = 0; $i -lt $Options.count; $i++) {
                        Write-Host -ForegroundColor Cyan "  $($i+1)." $Options[$i]
                    }
                    Write-Host # empty line
                    $global:ans = (Read-Host 'Please enter selection') -as [int]

                } While ((-not $ans) -or (0 -gt $ans) -or ($Options.Count -lt $ans))

                $global:selection = $Options[$ans - 1]

                #if AA
                if ($global:selection -eq "Teams Auto Attendant") {
                    Set-CsOnlineApplicationInstance -Identity $LCCRAS_ObjectID -ApplicationId $AAAID | Out-Null
                    #Sync-CsOnlineApplicationInstance -ObjectId $LCCRAS_ObjectID

                    $global:selection = $null
                    Clear-Host
                    if ($aas.count -le 1) {
                        $AA_Identity = $AAS
                        $AAName = $AAS.Name
                        Read-Host "Failing over to $AAName. Press ENTER to continue..."
                    }
                    if ($AAS.count -ge 2) {
                        Do {
                            Write-Host 'Failover to Teams Auto Attendant:'

                            for ($i = 0; $i -lt $AAS.count; $i++) {
                                Write-Host -ForegroundColor Cyan "  $($i+1)." $AAS.Name[$i]
                            }
                            Write-Host # empty line
                            $global:ans = (Read-Host 'Please enter selection') -as [int]

                        } While ((-not $ans) -or (0 -gt $ans) -or ($AAS.Count -lt $ans))

                        $global:selection = $AAS[$ans - 1]
                        $AA_Identity = $global:selection
                    }
                    New-CsOnlineApplicationInstanceAssociation -Identities @($LCCRAS_ObjectID) -ConfigurationId $AA_Identity.Identity -ConfigurationType AutoAttendant | Out-Null
                    Write-Host 'Please wait....'
                    start-sleep -Seconds 30
                    Write-Host 'Failover Auto Attendant config completed'
                }

                #if CQ
                if ($global:selection -eq "Teams Call Queue") {
                    Set-CsOnlineApplicationInstance -Identity $LCCRAS_ObjectID -ApplicationId $CQAID
                    #Sync-CsOnlineApplicationInstance -ObjectId $LCCRAS_ObjectID

                    $global:selection = $null
                    Clear-Host
                    if ($CQS.count -le 1) {
                        $CQ_Identity = $CQS
                        $CQName = $CQS.Name
                        Read-Host "Failing over to $CQName. Press ENTER to continue..."
                    }
                    if ($CQS.count -ge 2) {
                        Do {
                            Write-Host 'Fail over to Teams Call Queue:'

                            for ($i = 0; $i -lt $CQS.count; $i++) {
                                Write-Host -ForegroundColor Cyan "  $($i+1)." $CQS.Name[$i]
                            }
                            Write-Host # empty line
                            $global:ans = (Read-Host 'Please enter selection') -as [int]

                        } While ((-not $ans) -or (0 -gt $ans) -or ($CQS.Count -lt $ans))

                        $global:selection = $CQS[$ans - 1]
                        $CQ_Identity = $global:selection
                    }
                    New-CsOnlineApplicationInstanceAssociation -Identities @($LCCRAS_ObjectID) -ConfigurationId $CQ_Identity.Identity -ConfigurationType CallQueue | Out-Null
                    Write-Host 'Please wait....'
                    start-sleep -Seconds 30
                    Write-Host 'Failover Queue config completed'
                }
            }
        }
        catch { Throw "Cannot Failover to Teams. " + $error[0] }
        Try {
            if ($global:selection -eq "Failback Teams Queue or Auto Attendant to Landis Contact Center") {
                $global:selection = $null
                Clear-Host
                if ($ras.count -le 1) {
                    $RA_ObjectID = $RAS.ObjectId
                    $RAName = $RAS.DisplayName
                    Read-Host "Failing back $RAName. Press ENTER to continue..."
                }
                if ($RAS.count -ge 2) {
                    Do {
                        Write-Host 'Resource Account to failback to Landis Contact Center:'

                        for ($i = 0; $i -lt $RAS.count; $i++) {
                            Write-Host -ForegroundColor Cyan "  $($i+1)." $RAS.DisplayName[$i]
                        }
                        Write-Host # empty line
                        $global:ans = (Read-Host 'Please enter selection') -as [int]

                    } While ((-not $ans) -or (0 -gt $ans) -or ($RAS.Count -lt $ans))

                    $global:selection = $RAS[$ans - 1]
                    $RA_ObjectID = $global:selection.ObjectId
                }
                Remove-CsOnlineApplicationInstanceAssociation -Identities $RA_ObjectID | Out-Null
                Set-CsOnlineApplicationInstance -Identity $RA_ObjectID -ApplicationId $LCCAID | Out-Null
                #Sync-CsOnlineApplicationInstance -ObjectId $RA_ObjectID
                Write-Host 'Please wait....'
                start-sleep -Seconds 30
                Write-Host 'Failback config completed'
            }
        }
        catch { Throw "Cannot failback to Landis Contact Center. " + $error[0] }
        show-menu
    }
    show-menu
}