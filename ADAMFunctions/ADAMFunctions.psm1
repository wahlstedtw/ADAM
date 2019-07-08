#
# Functions for Automation ADAM 3.0
# Authors: Walter Wahlstedt
#
# v1.0|06/25/2019:	Functions for Automation.
#
# =====================================================================

# Make all errors terminating
$ErrorActionPreference = "Stop";

<# ====================================================================
Synopsis:
TODO: explain purpose
	Setup the symbolic link for the ADAMFunctions module in
		an elevated CMD prompt (PS doesn't work) from the desired profile.

	mkdir "%userprofile%\Documents\WindowsPowerShell\Modules"
	mklink /D "%userprofile%\Documents\WindowsPowerShell\Modules\ADAMFunctions" "{file share}\Account_Creation\ADAMFunctions"
=======================================================================
#>

<# ====================================================================

		Define Functions

========================================================================
#>

function Start-Countdown {
<# ====================================================================
	Initiates a countdown on your session.
	Usage:	Start-Countdown -Seconds 10 -ProgressBar
========================================================================
#>

Param(
[INT]$Seconds = (Read-Host "Enter seconds to countdown from"),
[Switch]$ProgressBar
)
	while ($seconds -ge 1){
		If($ProgressBar){
			Write-Progress -Activity "Countdown" -SecondsRemaining $Seconds -Status "Time Remaining"
			Start-Sleep -Seconds 1
		}ELSE{
			Write-Output $Seconds
			Start-Sleep -Seconds 1
		}
		$Seconds --
	}
}

Function Get-CurrentLine ( $addlines ) {
<# ====================================================================
	Gets the current line for debug
	Usage: Get-CurrentLine <number of lines to add>
========================================================================
#>
    $line = $Myinvocation.ScriptLineNumber
		$line = $line + $addlines
		return $line
}

function Set-LogFile ( $path, $name ){
<# ====================================================================
	Sets the log file name with the current date
========================================================================
#>
		$date = Get-date -Format 'yyyy-MM-dd'
		return "$path\Log\$name$($date).log"
}

function Start-LogRotate ( $path, $logname, $months, $days, $hours, $mins )
<# ====================================================================
	Rotates log files older than $months, $days, $hours, $mins
	Usage: Start-LogRotate <months> <days> <hours> <minutes>
========================================================================
#>
{
    $files = @(get-childitem $path | Where-Object {($_.LastWriteTime -lt (Get-Date).AddMonths(-$months).AddDays(-$days).AddHours(-$hours).AddMinutes(-$mins)) -and ($_.psIsContainer -eq $false) -and ($_.Name -like $logname)})
    if ($files -ne $NULL)
    {
        for ($idx = 0; $idx -lt $files.Length; $idx++)
        {
            $file = $files[$idx]
            Write-Host ("Rotate old log files to archive folder: " + $file.Name) -Fore Red
            Move-Item $path\$file $path\Archive
        }
    }
}

function Start-Logging ( $logPath, $logName, $months, $days, $hours, $mins, $logOption )
<# ====================================================================
	Starts the logging with start-transcript
	Rotates log files older than $months, $days, $hours, $mins
	Usage: Start-LogRotate <months> <days> <hours> <minutes>
========================================================================
#>
{
[hashtable]$return = @{}

$files = @(get-childitem $logPath | Where-Object {($_.LastWriteTime -lt (Get-Date).AddMonths(-$months).AddDays(-$days).AddHours(-$hours).AddMinutes(-$mins)) `
	-and ($_.psIsContainer -eq $false) -and ($_.Name -like $logName)})
if ($files -ne $NULL) {
	for ($idx = 0; $idx -lt $files.Length; $idx++){
		$file = $files[$idx]
		Write-Host ("Rotate old log files to archive folder: " + $file.logName) -Fore Red
		Move-Item $logPath\$file $logPath\Archive
	}
}
	$date = Get-date -Format 'yyyy-MM-dd'
	start-transcript "$logPath\Log\$logName$($date).log" #$logOption

	$return.Path = "$logPath\Log\"
	$return.Name = "$logName$($date).log"

Return $return
}

function Stop-Logging ()
<# ====================================================================
	Stops logging
	Usage: Stop-Logging
========================================================================
#>
{
	Stop-transcript
}


Function Get-RandomPassword ( $length ){
<# ====================================================================
	Used in the visitor password script, could probably merge the two.
	Generate a random password with at least
			1 Uppercase Letter
			1 Lowercase Letter
			1 Number
			1 Symbol
		Doesn't include 1,0,o,O,i,I,l,L,+
		48,49,73,76,79,105,108,111
	Usage: Get-RandomPassword <length>
========================================================================
#>
      $digits = 50..57
      $letters = 65..72 + 74..75 + 77..78 + 80..90 + 97..104 + 106..107 + 109..110 + 112..122
      $required = "!"

      while (!($password -cmatch "^(?=.*\p{Lu})(?:.*\p{Ll})(?=.*\d)(?=.*\W)(?!.*(.).*\1.*\1)")){

              $password = get-random -count $length `
                      -input ($symb + $digits + $letters) |
                              % -begin { $aa = $null } `
                              -process {$aa += [char]$_} `
                              -end {$aa}
              $password = $password + $required
      }

      return $password
  }


Function Remove-UserADGroups ( $username ){
<# ====================================================================
	Removes user from all groups and stores those groups in the notes property of the user
	Usage: Remove-UserADGroups <samaccountname>
========================================================================
#>
	Write-Output "`n Remove user groups for " $username

	# Check Group membership and populate the list to Notes
	$Groupmemberof += Get-ADUser -Identity $username -Properties memberof | ForEach-Object{
		$_.memberof | Get-ADGroup | Select-Object -ExpandProperty Name
	}

	Foreach ($Group in $GroupMemberof){
		#  Get the current notes and add the group membership from above
		$Notes = get-ADuser $username -Properties info | ForEach-Object{ $_.info }
		Set-ADUser $username -replace @{info="$($Notes) $($Group);"}
	}

	# Remove all memberships from the user except for "Domain Users"
	Get-ADUser $username -Properties MemberOf | ForEach-Object {
		$_.MemberOf  | where {$_.name -notmatch '^users|domain users$'} | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false
	}

}
Function Set-UserMailboxPermissions ( $user ){
<# ====================================================================
	Grants "Organization Management" FullAccess on a given mailbox
	Usage: Set-UserMailboxPermissions <user>
========================================================================
#>

		Add-MailboxPermission -Identity $user -User "Organization Management" -AccessRights FullAccess -InheritanceType all -AutoMapping:$false
	}
function Get-ScriptInfo {
<# ====================================================================
	Get script info for other functions returns script path and name variables
	Usage: $ScriptInfo = Get-ScriptInfo
		$ScriptInfo.Path
		$ScriptInfo.Name
========================================================================
#>
[hashtable]$return = @{}

		$return.Path = $MyInvocation.PSScriptRoot
		$return.Name = [io.path]::GetFileNameWithoutExtension($(split-path $MyInvocation.PSCommandPath -Leaf) )

	Return $return
}

Function Remove-UsersFromADGroup ( $groupName, $groupKeep){
	<# ====================================================================

	Removes all users from group

	========================================================================
	#>
			# $groupMembers = @()

write-host "Remove users from group"
$groupMembers = $groupName | Get-ADGroupMember | Where-Object {$_.name -NotMatch $groupKeep}

	foreach ($member in $groupMembers){
		Remove-ADGroupMember $groupName $member -Confirm:$false
				Write-Host $member
				Write-Host "Removed."
	}
}