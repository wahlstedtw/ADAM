#
# LocalAdmin Removal Script
# Authors: Walter Wahlstedt
#
# v1.0.0|WW|7/8/2019:		Initialization

<# ====================================================================
Synopsis:
  Sets the users of the LocalAdmins Group
=======================================================================
#>

Clear-Host
# Import AMS general functions and Active Directory modules.
try {
	Import-Module ActiveDirectory
	Import-Module ADAMFunctions -force
	# Get name and path of the current script.
	$ScriptInfo = Get-ScriptInfo
	$ScriptInfo.Path
	$ScriptInfo.Name
	# Records a log of the script output.
	$logInfo = Start-Logging $ScriptInfo.Path $ScriptInfo.Name 6 0 0 0 0
}
CATCH [system.exception]
{
	$FileName = [io.path]::GetFileNameWithoutExtension($(split-path $MyInvocation.InvocationName -Leaf))
	$FilePath = split-path $MyInvocation.InvocationName
		Start-Transcript "$FilePath\$FileName.log"
			Write-Output "Error loading modules, getting script info or starting the logging."
			$_.Exception.Message
		Stop-transcript
	exit
}
<# ====================================================================
	Variables
========================================================================
#>

$group = "{LocalAdmins}"
$keep = "{support admins}"
<# ====================================================================
	Start Main Function
========================================================================
#>
Write-output "Starting 	$ScriptInfo.Name"
Write-Host   "-----------------------------------------------------"
try {
    Remove-UsersFromADGroup $group $keep
    $success += ,@($group, $keep )

}
catch {
    Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
    $_.Exception.Message
    Write-Output "`n  -----------------------------------------------`n"
    $failed += ,@(($_.Exception.Message))
    $sendUserMail = $false
}

Write-Host   "-----------------------------------------------------"

Write-Output "AD Info Sync Finished."

# Send email to account creators with summary if anything was processed
# TODO: Good candidate for a function.
if ($sendAdminMail) {
	Write-Output "`n Send mail to admins."

	$adminBody  = "The following occurred when trying to set the Local Admin properties.<br />"
	$adminBody += "For a log of this transaction look here: <font color='blue'>$($logPath)</font><br /><br /><html><body>"

	# $adminBody += "For a log of this transaction look here: <font color='blue'>$($logInfo.Path)$($logInfo.Name)</font><br /><br /><html><body>"

	# Information about successful items.
	if ($success.count -gt 0){
		$adminBody += "<table border=1>"
		$adminBody += "<tr><th>Account Modifications</th><th>Password</th></tr>"
		foreach ($item in $success) { $priority = 'Normal'; $adminBody += "<tr><td> $($item[0]) </td> <td><font color='green'>$($item[1])</font></td></tr> " }
		$adminBody += "</table>"
	}
	# Information about failed items.
	if ($failed.count -gt 0){
		$adminBody += "<table border=1>"
		$adminBody += "<tr><th>Errors</th></tr>"
		foreach ($item in $failed) { $priority = 'High'; $adminBody += "<tr style='background-color: #D7978D;'><td> $($item[0]) </td></tr> " }
		$adminBody += "</table>"
	}
	# Information about the logs
	# $adminBody += " <br /><br />The following logs were moved to the archive folder: <br/><font color='blue'>$($logInfo.Path)\Archive</font><br/><br />"
	# $adminBody += "<table border=1>"
	# $adminBody += "<tr><th>Log Rotated to the archive folder.</th></tr>"
	# foreach ($item in $log) { $adminBody += "<tr style='background-color: #D7978D;'><td><font color='red'>$($item[0])</font></td></tr> " }
	# $adminBody += "</table>"
	$adminBody += "</body></html>"

	TRY
	{
		Send-MailMessage `
		-From "noreply-AMS@{yourdomain}" `
		-To "{your support team}" `
		-Subject "Account Update Summary" `
		-BodyAsHtml $adminBody `
		-SmtpServer "SMTP.domain.name" `
		-Priority $priority
	}
	CATCH [system.exception]
	{
		Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
		$_.Exception.Message
		Write-Output "`n  -------------------------------------------------------`n"

	}

}

Stop-Logging