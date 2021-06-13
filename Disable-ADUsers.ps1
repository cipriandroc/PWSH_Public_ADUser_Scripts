[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    $ticketNumber,
    [Parameter(Mandatory = $false)]
    [switch]$moveOU,
    [Parameter(Mandatory = $false)]
    [switch]$forceMoveOU
)

#region variables
[string]$ticketNumber = -join ('Disabled per AD Cleanup Project - ', $ticketNumber)
$importDataFile = -join ('.\', 'userlist', '.csv')
[string]$date = ((Get-Date).ToString() -split (' '))[0] -replace '/', '.'
[string]$exportLocation = 'C:\DATA\userDisable\'
[string]$exportBackupUserData = -join ($date, '_backupUserData', '.txt')
[string]$exportLog = -join ($date, '_disableReport', '.csv')
[string]$disableOU = 'OU=Disabled M365,OU=users,OU=_M365x565000.onmicrosoft.com,DC=orbi,DC=home'
[string]$targetDomain = 'orbi'

$VerbosePreference = "Continue"
#endregion variables
#region classes,functions
class FunctionReturnResults {
    [string]$message
    $value
    [int32]$returnCode = 1
}
class userObject {
    [string]$inputString
    [string]$samAccountName
    [string]$DistinguishedName
    [bool]$Enabled
    [string]$Description
    [string]$action = 'Could not disable'
    $movedOU = $null
}
function Write-CDOutputMessage {
    Param(
        [Parameter(Mandatory = $false)]
        [psobject]$Object,
        [Parameter(Mandatory = $false)]
        $greenifyMessage
    )
    
    if ($Object.returnCode -eq '0') { Write-Host $object.message -ForegroundColor 'Blue' -BackgroundColor 'DarkGreen' }
    elseif ($Object.returnCode -eq '1') { Write-Host $Object.message -BackgroundColor 'DarkRed' }
    if ($greenifyMessage) { Write-Host -Object $greenifyMessage -ForegroundColor 'Blue' -BackgroundColor 'DarkGreen' }
}
function Set-CDADAccountDescription {
    Param(
        [CmdLetBinding()]
        [Parameter(Mandatory = $true)]
        $user,
        [Parameter(Mandatory = $true)]
        $addInfo,
        $existingDescription
    )
    $functionReturnResults = [FunctionReturnResults]::new()

    if ($existingDescription) { $newDescription = -join ($ticketNumber, '|', $existingDescription) }
    else { $newDescription = $ticketNumber }
    Try {
        Set-ADUser -Identity $user -Description $newDescription -ErrorAction Stop -ErrorVariable changeDescrErr
        $functionReturnResults.message = -join ('Changed description to: ', $newDescription)
        $functionReturnResults.value = $newDescription
        $functionReturnResults.returnCode = 0
    }
    Catch {
        $functionReturnResults.message = -join ($user, ' Could not modify description: ', $changeDescrErr.message)
        $functionReturnResults.value = $changeDescrErr.message
    }
    return $functionReturnResults
}
function Disable-CDADAccount {
    Param(
        [CmdLetBinding()]
        [Parameter(Mandatory = $true)]
        $user,
        [Parameter(Mandatory = $true)]
        [bool]$enableStatus
    )

    $functionReturnResults = [FunctionReturnResults]::new()

    if (-not $enableStatus) {
        $functionReturnResults.message = -join ($user, ' is already disabled.')
        $functionReturnResults.value = 'already disabled'
        $functionReturnResults.returnCode = 0
    }
    else {
        Try {
            Disable-ADAccount -Identity $user -ErrorAction Stop -ErrorVariable disableFail
            $functionReturnResults.message = -join ($user, ' has been disabled.')
            $functionReturnResults.value = 'disabled'
            $functionReturnResults.returnCode = 0
        }
        Catch {
            $functionReturnResults.message = -join ($user, ' Could not disable: ', $disableFail.message)
            $functionReturnResults.value = $disableFail.message
        }
    }
    return $functionReturnResults
}
function Move-CDADObject {
    Param(
        [CmdLetBinding()]
        [Parameter(Mandatory = $true)]
        $userDataCollection,
        [Parameter(Mandatory = $true)]
        $disableOU
    )

    $functionReturnResults = [FunctionReturnResults]::new()

    Try {
        Move-ADObject -Identity $userDataCollection.DistinguishedName -TargetPath $disableOU -ErrorAction Stop -ErrorVariable failMove
        $functionReturnResults.message = -join ($userDataCollection.samAccountName, ' has been moved to disable OU')
        $functionReturnResults.value = 'moved'
        $functionReturnResults.returnCode = 0
    }
    Catch {
        $functionReturnResults.message = -join ($userDataCollection.samAccountName, ' ', $failMove.message)
        $functionReturnResults.value = $failMove.message
    }
    return $functionReturnResults
}
#endregion classes,functions
#region initialization
Try {
    Write-Verbose -Message ( -join ('Attempting to import userdata file ', $importDataFile)) -ErrorAction Stop -ErrorVariable failImport
    $userList = Import-Csv -Path $importDataFile
    Write-CDOutputMessage -greenifyMessage 'Succesfully imported userdata file.'
}
Catch { 
    Write-Warning -Message $failImport.message
    exit
}

Write-Verbose -Message 'Verifying user data file for samAccountName property.'
if (-not (( $userList | Get-Member | Where-Object { $_.MemberType -match 'NoteProperty' } | Select-Object -ExpandProperty Name ) -eq 'samAccountName')) {
    Write-Warning -Message "There's no 'samAccountName' property provided in the imported file"
    exit
}
Write-CDOutputMessage -greenifyMessage 'Property found.'

Write-Verbose -Message 'Verifying export folder location.'
if (-not (Test-Path $exportLocation) ) {
    Write-Warning -Message ( -join ('Could not find export folder ', $exportLocation))
    exit
}
Write-CDOutputMessage -greenifyMessage 'Test exportPath OK'

Write-Verbose -Message 'Verifying connected domain'
Try { 
    $runtimeDomain = (((Get-ADDomainController).domain).split('.'))[0] 
    if (-not ($runtimeDomain -eq $targetDomain) ) {
        Write-Warning -Message ( -join ('Runtime Domain', ' [', $runtimeDomain, '] ', 'different than target domain', ' [', $targetDomain, ']'))
        exit
    }
}
Catch {
    Write-Warning -Message 'Could not establish domain connection. Check network/vpn/credentials.'
    exit
}
Write-CDOutputMessage -greenifyMessage 'Domain connection OK'
if ($disableOU) {
    if ( -not ([adsi]::Exists( -join ('LDAP://', $disableOU)))) {
        Write-Warning -Message ( -join ('Could not find disable OU: ', $disableOU))
        exit
    }
}
#endregion initialization
#runspace
foreach ($user in $userList) {
    
    $createUserObject = [UserObject]::new()
    $createUserObject.inputString = $user.samAccountName
    $userData = $null

    if (-not $user.samAccountName) { 
        Write-Warning -Message ( -join ('No userdata provided in row: ', ($user.PSObject.Properties.Value -join ',')))
        $createUserObject.DistinguishedName = -join ('No userdata provided in row: ', ($user.PSObject.Properties.Value -join ','))
        $createUserObject.action = -join ('no valid entry')
    }
    else {
    
        Try { 
            Write-Verbose -Message ( -join ('Gathering user data for: ', $user.samAccountName))
            $userData = Get-ADUser -Identity $user.samAccountName -Properties * -ErrorAction Stop -ErrorVariable failFindUser 
        }
        Catch { 
            Write-Warning -Message $failFindUser.message 
            $createUserObject.DistinguishedName = $failFindUser.message 
            $createUserObject.action = 'not found'
        }

        if ($userData) {
            Write-Verbose -Message ( -join ('Backing up user data to file: ', $exportBackupUserData))
            $userData | Out-File -FilePath ( -join ($exportLocation, $exportBackupUserData)) -Append

            if ($userData.Enabled) {
                Write-Verbose -Message ( -join ('Attempting to change description filed for user: ', $userdata.samAccountName))
                $changeDescription = Set-CDADAccountDescription -user $userData.samAccountName -addInfo $ticketNumber -existingDescription $userData.Description
                Write-CDOutputMessage -Object $changeDescription

                Write-Verbose -Message ( -join ('Attempting to disable user: ', $userdata.samAccountName))
                $disableAccount = Disable-CDADAccount -user $userData.samAccountName -enableStatus $userData.Enabled
                Write-CDOutputMessage -Object $disableAccount
                if ($disableAccount.returnCode -eq 0) {
                    $createUserObject.action = $disableAccount.value
                }
            }
            else {
                Write-CDOutputMessage -greenifyMessage ( -join ('User is already disabled: ', $userData.samAccountName))
                $createUserObject.action = 'already disabled'
            }
            if ($moveOU) {
                if ($userData.Enabled -or $forceMoveOU) {
                    Write-Verbose -Message 'Attempting to move user to disable OU'
                    $moveUserOU = Move-CDADObject -userDataCollection $userData -disableOU $disableOU
                    Write-CDOutputMessage -Object $moveUserOU
                    if ($moveUserOU.returnCode -eq 0) {
                        $createUserObject.movedOU = $true
                    }
                    elseif ($moveUserOU.returnCode -eq 1) {
                        $createUserObject.movedOU = $false
                    }
                }
            }

            $verifyUserData = Get-ADUser -Identity $user.samAccountName -Properties Enabled, Description

            foreach ($property in $createUserObject.PSObject.Properties.Name) {
                if (($property -eq 'action') -or ($property -eq 'inputString') -or ($property -eq 'movedOU')) {
                    continue
                }
                $createUserObject.$property = $verifyUserData.$property
            }
        }
    }

    Write-Verbose -Message 'Exporting disable information to file'
    $createUserObject | Export-Csv -Path ( -join ($exportLocation, $exportLog)) -Append -NoTypeInformation
}
#endrunspace

#todo:
#ticket number validation (contains hd-43242? numbers , dash)




