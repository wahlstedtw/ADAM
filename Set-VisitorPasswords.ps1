#
# Visitor Account Password Reset
# Author: Walter Wahlstedt
# August 6, 2013
# Version 1.0: Resets Visitor account passwords when ran and emails them to a Distribution Group
# Version 1.1: Adds function file
#
# ADAMFunctions module is a symbolic link from the file share\account management\ADAMFunctions
# folder to the C:\Users\powershellservice\Documents\WindowsPowerShell\Modules folder

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


$emailFile = "{file share}\account_creation\Mail_Templates\EmailVisitorPass.htm"
$sendUserMail = $true
$sendAdminMail = $true
# Default is 5 visitor accounts which are already in Active Directory.
$numVisitorAccounts = 5
$visitorPW
$failed = @()
$success = @()
$adminBody = @()

<# ====================================================================

 Start Main Function

========================================================================
#>
write-host "Reset Visitor Account Passwords"
TRY{
# Loop through the num of visitor accounts and reset the passwords
for ($i =1 ; $i -le $numVisitorAccounts; $i++) {
	# Reset the visitor password
    Set-ADAccountPassword -Identity visitor$i -Reset -NewPassword (ConvertTo-SecureString -AsPlainText ($visitorPW = Get-RandomPassword 8) -Force)
    $success += ,@(("visitor" + $i), $visitorPW )
}
}Catch{
	Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
		$_.Exception.Message
	Write-Output "`n  -----------------------------------------------`n"
	$failed += ,@(($_.Exception.Message))
	$sendUserMail = $false
}

if ($sendUserMail) {
Write-host "Send e-mail to distribution group"

$userBody = [string]::join([environment]::newline, (Get-Content -path $emailFile -Raw))
	foreach ($item in $success) {
		$userBody = $userBody.Replace("[$($item[0])]", $item[1])
		write-host $($item[0]) : $($item[1])
	}
# This is where we compose and send the email
# TODO: Good candidate for a function.
	Send-MailMessage `
	-From "{Your support account}" `
	-To "{Users who distribute visitor accounts}" `
	-Subject "Visitor Password Summary" `
	-BodyAsHtml $userBody `
	-SmtpServer "SMTP.domain.name" `
	-Priority $priority
}
# Send email to account creators with summary if anything was processed
# TODO: Good candidate for a function.
if ($sendAdminMail) {
    Write-Output "`n Send mail to admins."

	$adminBody  = "The following occurred when trying to set the visitor passwords.<br />"
	$adminBody += "For a log of this transaction look here: <font color='blue'>$($logPath)</font><br /><br /><html><body>"
	# $adminBody += "For a log of this transaction look here: <font color='blue'>$($logInfo.Path)$($logInfo.Name)</font><br /><br /><html><body>"

	# Information about successful items.
    if ($success.count -gt 0){
        $adminBody += "<table border=1>"
        $adminBody += "<tr><th>Visitor Action</th><th>Password</th></tr>"
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