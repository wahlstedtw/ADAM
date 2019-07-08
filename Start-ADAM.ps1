#
# ADAM (Active Directory Account Management)(AMS is the old name)
# Authors: Nolan Davidson, Walter Wahlstedt, Jason Frazier
#
# v1.0|ND|06/09/2011:		Account creation script.
# v1.1|ND|09/26/2011:		Added emailing the students their information
# v2.0|JF|05/06/2013:		Added Exchange mailbox creation
#													Implemented new functions: email, sync, SetSecurity
# v2.1|WW|06/12/2013:		Implemented new functions: update, notify, disable,
#													delete, Strip
# v2.2|WW|07/25/2013:		Fixes the Update Function.
# v2.3|WW|05/23/2016:		All emails are sent at the end of the script.
#													Added Set-Logfile function for transcripts.
#													Added date to log file
#													Log file now appends if there are two logs with the same name
#													Delete now only deletes users from pending removals
#													Added Alumni action
#													relocated email and security files
#													Added chars to password generation to force
#													compliance with ad password policy
# v2.4|WW|06/15/2016:		Upgraded to AD Connect.
#													all errors are terminating for try catch
#													staff, faculty and adjunct are now created in office365
#													create fac or staff drives
#													turn clutter off
# v2.5|WW|01/18/2017:		Renamed Functions
#													Added logrotate Function
#													Added logrotate details to admin email
# v2.6|WW|02/17/2017:		added Set-UserMailboxPermissions function
# v2.7|WW|06/14/2018:		Migrated from app05 to AADSync01
#													Fixed Logfile link in email
#													Changed error text color to black
#													added Get-CurrentLine function
#													Removed Unnecessary code for local exchange create
# v2.8|WW|07/26/2018:		Set Service for powershell remote connections.
#													Fixed try/catch terminating errors
#													Adjusted wording
#													created $seconds variable
# v2.9|WW|08/2/2018:		additional error checking for AAD sync.
#													Added line breaks for all Write-Output commands
# v2.9.1|WW|08/15/2018:	Removed partial dependency on Quest tools.
#													Reworked the Disable Action.
#													Reworked the Remove-UserADGroups function.
# v3.0.1|WW|07/2/2019:	Moved functions to module ADAMFunctions
# =====================================================================

# Make all errors terminating
$ErrorActionPreference = "Stop";

<# ====================================================================
Synopsis:
	Create
		Will create the user account, set mailbox properties, set random password, set distribution group, OU and Custom attributes for PWM. For 365 users it will assign them a license as well.
	Update
		Changes the users, Email, first name, last name, middle initial, display name, username. Then emails the user.
	Notify
		Removes user from all groups and stores those groups in the notes property of the user then emails the user.
	Disable
		Disables the user account, moves them to pending removals and emails the user.
	Delete
		Deletes the user account from pending removals.
	Sync
		Runs a Directory Sync
	Email
		Emails the specified account.
	SetSecurity
		allows you to create new security files for exchange and o365
	Strip
		Removes user from groups and adds those groups to account notes
	Alumni
		Moves a user from Pending removals to Alumni OU and adds them to the alumni group
=======================================================================
#>

Clear-Host
# Import AMS general functions and Active Directory modules.
try {
	Import-Module ActiveDirectory
	Import-Module ADAMFunctions -force
	# TODO: Remove reliance on quest tools
	Add-PSSnapin Quest.ActiveRoles.ADManagement
	# Get name and path of the current script.
	$ScriptInfo = Get-ScriptInfo
	$ScriptInfo.Path
	$ScriptInfo.Name
	# Start recording a log of the script output.
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


Clear-Host

$logPath = "{path\to\log\}Log"
Start-Transcript (Set-Logfile $logPath) -Append

# Config variables
$priority = 'low'
$parentOU = ''
$year = ''
$baseGroup = ''
$sync = $false
$moveToPendingRemovals = '{inactive users ou}'
$moveToAlumni = '{Alumni users ou}'
$seconds = 30


# Import csv
$csv = Import-Csv ("{path the account management csv}DataFiles\account_management.csv")
$O365accounts = @()

# Arrays to hold successful and unsuccessful creations.
$success = @()
$successCount = $null
$failed = @()
$failedCount = $null
$queueMail = @()

$exchangeCnt = $true
$sendUserMail = $false
$sendAdminMail = $true


# Array for OU Testing
$OU = @('OU={Sub OU 1},OU={Sub OU}','OU={Sub OU 2},OU={Sub OU}')
$RootOU = 'OU={Authorized Users},DC={your},DC={domain}'

# Test to make sure parent OU's exist
$OUexists = $true
foreach ($PathOU in $OU){
    $Path = "$PathOU,$RootOU"
    if (!([adsi]::Exists("LDAP://$Path"))) {
            $failed += ,@($line.account_type,$line.username, $line.id, "Supplied Path does not exist:  `n$Path", $line.account_type) # not going to populate fields because csv hasnt been read yet.
            # Don't try to process any accounts, Send email to admins
            $OUexists = $false
            break
        }
}

<# ====================================================================

	Start Script

========================================================================
#>

foreach ($line in $csv){
	$sendUserMail = $false
	$sendAdminMail = $true
	# Make sure the OU exists
	if (!$OUexists) {
		$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Supplied Path does not exist: " + $Path))
        break
    }

	switch ($line.account_action) {
		CREATE {
			Write-Output "`n ----------------------------- Create Action -----------------------------`n"

			if ($line.account_type -eq "none") {
				$sendAdminMail = $false
				break
			}

			# Make sure the account type code is valid
			if (($line.account_type -ne "staff") -and ($line.account_type -ne "faculty") -and ($line.account_type -ne "grad") -and ($line.account_type -ne "undergrad") -and ($line.account_type -ne "adjunct") -and ($line.account_type -ne "physicalplant") -and ($line.account_type -ne "foodservice")) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Invalid Code: " + $line.account_type))
				Write-Output Invalid Code
			}
			# Make sure username isn't already in use.
			elseif ($user = Get-QADUser $line.username) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Duplicate Username: " + $user.samAccountName))
				Write-Output "Username in use"
			}
			# Make sure Display Name isn't already in use.
			elseif ($user = Get-QADUser $line.display_name) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Duplicate Display Name: " + $line.display_name))
				Write-Output "Display Name in use"
			}
			# Make sure the student/employee ID doesn't already have an account.
			elseif ($user = Get-QADUser -ObjectAttributes @{extensionAttribute1=$line.id}) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Duplicate ID: " + $user.description))
				Write-Output "ID already has account"
			}

			# Check for Account type
			else {
				$code = $line.account_type
				$year = $line.year

					if ($code -eq "grad") {
						$parentOU = "{OU distinguished name}"
						$baseGroup = "students_grad"
						$exchDB = $null
					} elseif ($code -eq "undergrad") {
						$parentOU = "{OU distinguished name}"
						$baseGroup = "students_undergrad"
						$exchDB = $null
					} elseif ($code -eq "staff") {
						$parentOU =  '{OU distinguished name}'
						$baseGroup = "staff"
						$exchDB = "staff"
					} elseif ($code -eq "faculty") {
						$parentOU =  '{OU distinguished name}'
						$baseGroup = "Faculty"
						$exchDB = "faculty"
					}elseif ($code -eq "adjunct") {
						$parentOU = '{OU distinguished name}'
						$baseGroup = "" # no base group
						$exchDB = "faculty"
					}elseif ($code -eq "physicalplant") {
						$parentOU = '{OU distinguished name}'
						$baseGroup = "" # no base group
					}elseif ($code -eq "foodservice") {
						$parentOU = '{OU distinguished name}'
						$baseGroup = "" # no base group
					}
				 # Generate Random Password
				$passwd = Get-RandomPassword 25
				$SecurePassword = $passwd | ConvertTo-SecureString -AsPlainText -Force

				if($exchangeCnt){
				Write-Output "Connect to local exchange for create."
				$username = "{domain\username}"
				$password = cat "{path to credentials}\cred.txt" | convertto-securestring
				$exchangeConnection = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password
					$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://{exchange}/PowerShell/ -Authentication Kerberos -Credential $exchangeConnection

					TRY{
				 		Import-PSSession $exchangeSession -DisableNameChecking
					}Catch{
						Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
						$_.Exception.Message
						Write-Output "`n  -----------------------------------------------`n"
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					$exchangeCnt = $false
				}

				if ($code -eq "notbeingused"){ # put here it ignore creating on on-premise accounts

					if ($exchDB){
					Write-host "Create on-premise"

						TRY{
						$null = New-mailbox `
							-Name $line.display_name `
							-DisplayName $line.display_name `
							-UserPrincipalName $line.email `
							-samAccountName $line.username `
							-Firstname $line.first_name `
							-Lastname $line.last_name `
							-initials $line.middle_initial `
							-Password $SecurePassword `
							-OrganizationalUnit $parentOU
						}Catch{
							write-host  ----------------- new mailbox Exception --------------------`n
							$_.Exception.Message
							write-host  ------------------------------------------------`n
							$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
						}
						TRY{
						Set-Mailbox `
							-Identity $line.email `
							-CustomAttribute1 $line.id `
							-CustomAttribute2 $line.ssn_last_4 `
							-CustomAttribute3 $line.birth_date `
							-CustomAttribute5 ($line.id + 'uc')
						}Catch{
							write-host  ----------------- set-mailbox Exception --------------------`n
							$_.Exception.Message
							write-host  ------------------------------------------------`n
							$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
						}
						TRY{
						Add-DistributionGroupMember `
							-Identity $code `
							-Member $line.email `
							-BypassSecurityGroupManagerCheck
							$location = "On-Premise"
						}Catch{
							write-host  ----------------- add-distribution Exception --------------------`n
							$_.Exception.Message
							write-host  ------------------------------------------------`n
							$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
						}
						TRY{
						Set-aduser `
							-Identity $line.username `
							-homeDirectory ("{unc path}\" + $line.username)`
							-homeDrive Z:
						}Catch{
							write-host  ----------------- set-aduser Exception --------------------`n
							$_.Exception.Message
							write-host  ------------------------------------------------`n
							$failed += ,@($line.account_action, $line.account_type, $line.username, $line.id, $_.InvocationInfo.ScriptLineNumber, ($_.Exception.Message))
						}
					}else{
						write-host exchDB is null
						$location = "Error - No exchange database specified - line 169"
					}

				}elseif ($code){

					Write-Output "Create User in Office 365."
					if( ($code -eq "grad") ){
						$userCode = "students_grad"
					}
					elseif( ($code -eq "undergrad") ){
						$userCode = "students_undergrad"
					}
					else{ $userCode = $code }

					$rutAddr = $line.username + "@{0365 domain}.mail.onmicrosoft.com"

					TRY{
					$null = New-RemoteMailbox `
						-Name $line.display_name `
						-DisplayName $line.display_name `
						-UserPrincipalName $line.email `
						-samAccountName $line.username `
						-Firstname $line.first_name `
						-Lastname $line.last_name `
						-initials $line.middle_initial `
						-Password $SecurePassword `
						-OnPremisesOrganizationalUnit $parentOU `
					}Catch{
						Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
						$_.Exception.Message
						Write-Output "`n  -----------------------------------------------`n"
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					TRY{
					Set-RemoteMailbox `
						-Identity $line.email `
						-CustomAttribute1 $line.id `
						-CustomAttribute2 $line.ssn_last_4 `
						-CustomAttribute3 $line.birth_date `
						-CustomAttribute5 ($line.id + 'uc')
					}Catch{
						Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
						$_.Exception.Message
						Write-Output "`n  -----------------------------------------------`n"
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					TRY{
					Add-DistributionGroupMember `
							-Identity $userCode `
							-Member $line.email `
							-BypassSecurityGroupManagerCheck
					}Catch{
						Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
						$_.Exception.Message
						Write-Output "`n  -----------------------------------------------`n"

						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					$O365accounts += ,@{
							Code = $code
							UserPrincipalName = $line.email
					}
					$location = "Office 365"
					$sync = $true
				}

				Write-Output "Account " $line.username " created"

				$emailFile = "{path to email template}\EmailCreatedUsers.htm"
				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, 'Successfully Created' + $location)
				$sendUserMail = $true
				$emailToUC = $false
			}
			Write-Output "`n ------------------------------------------------------------------------`n"
		}
		UPDATE{
			Write-Output "`n ----------------------------- Update Action ------------------------------`n"
			$code = $line.account_type
			if ($exchangeCnt){
			# adminexchange.txt
				Write-Output "Connect to exchange for update"
				$username = "{domain\username}"
				$password = cat "{path to credentials}\cred.txt" | convertto-securestring
				$exchangeConnection = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password
				$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange/PowerShell/ -Authentication Kerberos -Credential $exchangeConnection
				TRY{
					Import-PSSession $exchangeSession -DisableNameChecking
			 	}Catch{
					Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
					$_.Exception.Message
					Write-Output "`n  -----------------------------------------------`n"
					$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
			 	}
				$exchangeCnt = $false
			}

			Write-Output "Attempt to change the account from " $line.display_name " to " $line.new_display_name
			if (Get-Mailbox $line.email){
				if ((Get-QADUser $line.username) -and ($line.new_first_name) -and ($line.new_last_name) -and ($line.new_display_name) -and ($line.new_email) -and ($line.new_username)) {
					Try{
					Set-User `
						-Identity $line.email `
						-Firstname $line.new_first_name `
						-Lastname $line.new_last_name `
						-initials $line.new_middle_initial `
						-name $line.new_display_name

					Set-Mailbox `
						-Identity $line.email `
						-EmailAddressPolicyEnabled $false `
						-PrimarySMTPAddress $line.new_email `
						-DisplayName $line.new_display_name `
						-UserPrincipalName $line.new_email `
						-samAccountName $line.new_username `
						-alias $line.new_username

					Set-Mailbox -Identity $line.new_email -EmailAddressPolicyEnabled $true
					}
					catch{
						Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
						$_.Exception.Message
						Write-Output "`n  -----------------------------------------------`n"
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					Try{
						Rename-Item -path ("{path to home directory}" + $line.username) -newName $line.new_username
					}
					catch{
						Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
						$_.Exception.Message
						Write-Output "`n  -----------------------------------------------`n"
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}

					$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("User changed from " + $line.username + " to " + $line.new_username) )
					$sendUserMail = $true
					$emailToUC = $true
					$emailFile = "{path to email template}\EmailUpdateUsers.htm"
				}
				else{
					Write-Output "`n New property is blank"
					$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("New user property is blank"))
				}
			}elseif (Get-RemoteMailbox $line.email){
					Write-Output "Update User Office 365"
					if( ($code -eq "grad") ){
						$userCode = "students_grad"
					}
					elseif( ($code -eq "undergrad") ){
						$userCode = "students_undergrad"
					}
					else{ $userCode = $code }

					$rutAddr = $line.username + "@{0365 domain}.mail.onmicrosoft.com"
					Try{
						Set-User `
							-Identity $line.email `
							-Firstname $line.new_first_name `
							-Lastname $line.new_last_name `
							-initials $line.new_middle_initial `
							-name $line.new_display_name

						Set-Mailbox `
							-Identity $line.email `
							-EmailAddressPolicyEnabled $false `
							-PrimarySMTPAddress $line.new_email `
							-DisplayName $line.new_display_name `
							-UserPrincipalName $line.new_email `
							-samAccountName $line.new_username `
							-alias $line.new_username

						Set-Mailbox -Identity $line.new_email -EmailAddressPolicyEnabled $true

						Add-DistributionGroupMember `
							-Identity $userCode `
							-Member $line.new_email `
							-BypassSecurityGroupManagerCheck

						#$O365accounts += ,@($code, $line.email)
						$location = "Office 365"
						$sync = $true
						$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("User changed from " + $line.username + " to " + $line.new_username) )
						$sendUserMail = $true
						$emailToUC = $true
						$rutAddr = $line.username + "@{0365 domain}.mail.onmicrosoft.com"
					}
						catch{
							Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
							$_.Exception.Message
							Write-Output "`n  -----------------------------------------------`n"
							$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
						}


			}else{
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Unknown Issue"))
			}
			Write-Output "`n ------------------------------------------------------------------------`n"
		}
		NOTIFY {
			Write-Output "`n ---------------------------- Notify Action ------------------------------`n"

			# Make sure username exists

			if (!($user = Get-QADUser $line.username)) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $user.samAccountName) )

				Write-Output "Account " $line.username " doesn't exist"
			}
			else{
				Remove-UserADGroups $line.username

				$emailFile = "{path the email template}\EmailNotifyUsers.htm"

				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Removed permissions and notified " + $line.personal_email + ". Disable Pending") )
				Write-Output ("Notify Account: " + $line.username)
				$sendUserMail = $true
				$emailToUC = $true
			}
			Write-Output "`n --------------------------------------------------------------------------`n"
		 }
		DISABLE {
			Write-Output "`n ---------------------------- Disable Action ------------------------------`n"
			$sam = $line.username
			if (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}')) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $line.samAccountName))
				Write-Output "Account " $line.username " doesn't exist"
				$sendUserMail = $false
			}
			elseif (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'OU={Pending Removals},DC={domain},DC={name}')){
				# Strip the groups from the user and add to notes
				Remove-UserADGroups $line.username

				Try{
				# Disable User account
				# Use $null to suppress unnecessary info in log
				$null = Disable-ADAccount $line.username

				}CATCH [system.exception]
				{
					Write-Output "`n  ---------------- Exception --------------------`n"
					$_.Exception.Message
					Write-Output "`n  -----------------------------------------------`n"
				  $failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
				}

				# move this user
				TRY{
				Write-Output "`n Move User - " $line.username
				get-aduser $line.username | Move-ADObject -TargetPath $moveToPendingRemovals
				Write-Output "`n Moved"
				}CATCH [system.exception]
				{
					Write-Output "`n  ---------------- Exception --------------------`n"
					$_.Exception.Message
					Write-Output "`n  -----------------------------------------------`n"
				   $failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
				}


				$emailFile = "{path the email template}\EmailDisableUsers.htm"

				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, 'Account disabled and moved to pending removals. Deletion pending.')
				Write-Output "Account Disabled"
				$sendUserMail = $false
				$emailToUC = $false
			}
			else{
				# Strip the groups from the user and add to notes
				Remove-UserADGroups $line.username

				Try{
					# Disable User account
					# Use $null to suppress unnecessary info in log
					$null = Disable-ADAccount $line.username
					Write-Output "`n Disabling Account already in Pending Removals."

				}CATCH [system.exception]
				{
					Write-Output "`n  ---------------- Exception --------------------`n"
					$_.Exception.Message
					Write-Output "`n  -----------------------------------------------`n"
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
				}

				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account stripped, disabled, already in pending removals."))
				Write-Output $line.username " Account has already been stripped, disabled and moved to pending removals."
				}
			Write-Output "`n ------------------------------------------------------------------------`n"
		 }
		DELETE {
			Write-Output "`n ---------------------------- Delete Action ------------------------------`n"
			$sam = $line.username

			# Deletes users only from Pending Removals
			if ($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'OU={Pending Removals},DC={domain},DC={name}'){

				TRY{

				Write-Output "`n attempting to delete " $sam
				Remove-ADObject $user -Recursive -Confirm:$false
				Write-Output "Account " $line.username " Deleted"
				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, 'Account Deleted.')

					if (($line.account_type -eq "undergrad") -or ($line.account_type -eq "grad")){
						$sync = $true
					}
				}CATCH [system.exception]
				{
					Write-Output "`n  ---------------- Exception --------------------`n"
					$_.Exception.Message
					Write-Output "`n  -----------------------------------------------`n"
					$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
				}

			}
			elseif ($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}') {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($user.samAccountName + " is not in pending removals. LINE:"+ (Get-CurrentLine)))
				Write-Output "Account is not in pending removals."
			}
			else{
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $line.username))
				Write-Output "Account " $line.username " doesn't exist"
			}
			Write-Output "`n ------------------------------------------------------------------------`n"
		 }
		SYNC{
			Write-Output "`n ---------------------------- Sync Action ------------------------------`n"
			 $sync = $true
			 $seconds = 0
			 $success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Syncing AD to Azure AD") )
			Write-Output "`n ------------------------------------------------------------------------`n"
		}
		EMAIL{
			Write-Output "`n  ---------------------------- Send Test Mail Action ------------------------------`n"
			$sendUserMail = $true
			$emailFile = "{path the email template}\EmailCreateduers.htm"
			$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Send Test Email: " + $line.personal_email) )
			Write-Output "`n ------------------------------------------------------------------------`n"
		}
		SETSECURITY{
			Write-Output "`n Set Security"
			Write-Output "This Sets the Security File for both the Local and Cloud AMSservice accounts."

			$AMS_Ver = read-host -prompt "Account Management Version (lowercase)"
		 	read-host -prompt "Input Password for {AMS Service Account}"  -assecurestring |convertfrom-securestring | out-file ("{UNC Path to }\Account_Creation\Secure_Files\" + $AMS_Ver + "\local-cred.txt")
			read-host -prompt "Input Password for {office 365 admin account}"  -assecurestring |convertfrom-securestring | out-file ("{UNC Path to }\Account_Creation\Secure_Files\" + $AMS_Ver + "\cloud-cred.txt")
			$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Reset Security Strings"))
		}
		STRIP{
			Write-Output "`n ---------------------------- Strip Action ------------------------------`n"
			$sam = $line.username
			if ($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}'){
				Remove-UserADGroups $line.username
				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Stripped Permissions - " + $line.username) )
			}
			else{
				Write-Output "Account " $line.username " doesn't exist"
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist"))
				}
			Write-Output "`n ------------------------------------------------------------------------`n"
		}
		ALUMNI{
			Write-Output "`n ---------------------------- Alumni Action ------------------------------`n"
			$sam = $line.username
			$sendUserMail = $false
			# Check to see if user account exists
			if (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}')) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $line.samAccountName))
				Write-Output "Account " $line.username " doesn't exist"

			}
			# Check to see if account is in pending removals then move to the alumni OU and add them to alumni group
			elseif (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'OU={Pending Removals},DC={domain},DC={name}')){
				# Remove User from groups
				Write-Output "`n remove user from groups"
				Remove-UserADGroups $line.username
				Write-Output "`n done"
				# move this user
				Write-Output "`n Move User to alumni - " $line.username
				get-aduser $line.username | Move-ADObject -TargetPath $moveToAlumni
				Write-Output 'User Moved'

				if (Get-Mailbox $line.email){
					Add-DistributionGroupMember `
						-Identity alumni `
						-Member $line.email `
						-BypassSecurityGroupManagerCheck

					Write-Output done
					$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Permissions Removed and moved to Authorized Users/Alumni Users") )
				}
				else{
					Write-Output "`n New property is blank"
					$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Can't find mailbox"))
				}
			}
			Write-Output "`n ------------------------------------------------------------------------`n"
		}
		default {
			if (($line.account_action -eq "none") -or ($line.account_action -eq $null)) {
				Write-Output "`n  ---------------------------- No Action ------------------------------`n"
				$sendAdminMail = $false
				Write-Output "`n ------------------------------------------------------------------------`n"
				break
			}
			else{
				Write-Output ("NO match for the --  " + $line.account_action + "  --action.")
				$sendAdminMail = $true
			}
		}
	}

	if ($sendUserMail) {

		$userBody = [string]::join([environment]::newline, (cat -path $emailFile))
		$userBody = $userBody.Replace('[userName]', $line.username)
		$userBody = $userBody.Replace('[firstName]', $line.first_name)
		$userBody = $userBody.Replace('[idNumber]', $line.ID)
		$userBody = $userBody.Replace('[new_userName]', $line.new_username)
		$userBody = $userBody.Replace('[new_displayName]', $line.new_display_name)

		if ($emailToUC) {
			$userTo = ($line.personal_email, $line.email)
			}
		else {
			$userTo = $line.personal_email
			}
		Write-Output "Queue e-mail to " $userTo
		$queueMail += , @($userTo,$userBody)
	}
}

if(!$exchangeCnt){Remove-PSSession $exchangeSession}

if ($sync)
{
	Write-Output "`n sync"
	# Provide username and securestring to AADSync Server for Remote powershell
	Write-Output "`n Waiting for local active directory to sync before running Azure Sync"

	Start-Countdown -Seconds $seconds


	# Creds for connection to AADSync Server
	Write-Output "Connect to {AADSync Server} for sync"
	$username = "{domain\AMS Service}"
	$password = cat "{Path to secure files}\Account_Creation\Secure_Files\prod\local-cred.txt" | convertto-securestring
	$dirsync = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password

	# Start DirSync on AADSync Server
	TRY
	{
		Write-Output "`n Try to Run Directory Sync on AADSync Server"
		Invoke-Command -ComputerName "{AADSync Server}" -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } -Credential $dirsync
	 $seconds = 120

	}
	CATCH [system.exception]
	{
		Write-Output "`n  ----------------  Exception  : Line $(Get-CurrentLine) --------------------`n"
		$_.Exception.Message
		Write-Output "`n  -----------------------------------------------`n"

		 TRY{
			 		Write-Output "`n Trying to Run Directory Sync on AADSync Server, something went wrong last time"
					Invoke-Command -ComputerName "{AADSync Server}" -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } -Credential $dirsync

		 }
		 	CATCH [system.exception]
			{
				Write-Output "`n  ----------------  Exception  : Line $(Get-CurrentLine) --------------------`n"
				$_.Exception.Message
				Write-Output "`n  -----------------------------------------------`n"
				# Queue email to admins
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message + ": Line " + $(Get-CurrentLine)))
				$seconds = 0

				$O365accounts = $null
				$queueMail = $null
			}
	}

	# Delay for DirSync to finish running
	Start-Countdown -Seconds $seconds
}

if ($O365accounts)
{
	TRY
	{
		Write-Output "`n O365 Accounts"
		#Connect to Office 365 Remote Powershell
		$pass = cat "{path the secure files}Prod\cloud-cred.txt" | convertto-securestring
		$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist "{o365 admin email}",$pass

		$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/  -Authentication Basic -AllowRedirection -Credential $cred
	}
	CATCH [system.exception]
	{
		Write-Output "`n  ----------------  Exception  : Line $(Get-CurrentLine) --------------------`n"
		$_.Exception.Message
		$_.Exception.ItemName
		Write-Output "`n  -----------------------------------------------`n"

		$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message + ": Line " + (Get-CurrentLine)))
	}
	TRY
	{
		Write-Output "`n Try to import PSsession, module and connect to MsolService for O365"
		Import-PSSession $O365Session -DisableNameChecking | out-null
		Import-Module MSOnline
		Connect-MsolService -Credential $cred
	}
	CATCH [system.exception]
	{
		Write-Output "`n  ----------------  Exception  : Line $(Get-CurrentLine) --------------------`n"
		$_.Exception.Message
		$_.Exception.ItemName
		Write-Output "`n  -----------------------------------------------`n"

		$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message + ": Line " + (Get-CurrentLine)))
	}

	Write-Output "`n Try to Assign License."
	#Loop through csv to assign user license
	foreach ($account in $O365accounts)
	{
		$account.code
		$account.UserPrincipalName
		Write-Output $account.code
		if ($account.code -eq "none")
		{
			Write-Output "`n A Code of none was found."
			break
		}
		if (($account.code -eq "grad") -or ($account.code -eq "undergrad"))
		{
			$license = 'unionkyedu:STANDARDWOFFPACK_IW_STUDENT'
			Write-Output "`n A " $account.code " license " $license " will be assigned to " $account.UserPrincipalName
		}
		elseif ($account.code)
		{
			$license = 'unionkyedu:STANDARDWOFFPACK_IW_FACULTY'
			Write-Output "`n A " $account.code " license " $license " will be assigned to " $account.UserPrincipalName
		}

		TRY
		{
			Write-Output "`n Try to set license O365."
			Set-Msoluser -UserPrincipalName $account.UserPrincipalName -UsageLocation US
			Set-MSOLUserLicense -UserPrincipalName $account.UserPrincipalName -AddLicenses $license
			Write-Output ( $account.code + " license assigned to user " + $account.UserPrincipalName)
		}
		CATCH [system.exception]
		{
			Write-Output "`n  ----------------  Exception  : Line $(Get-CurrentLine)  --------------------`n"
			$_.Exception.Message
			Write-Output "`n  -----------------------------------------------`n"

			$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message + ": Line " + $(Get-CurrentLine)))
		}
	}
	Start-Countdown -Seconds $seconds
	Write-Output "`n Set Mailbox properties."
	foreach ($account in $O365accounts){
		Write-Output "`n testing office365 accounts array: " $account "`n"
	}

	foreach ($account in $O365accounts)
	{
		TRY
		{
			Write-Output "`n Verify exchange server functionality for " $account.UserPrincipalName "`n"
			$null = Test-MapiConnectivity $account.UserPrincipalName

			Write-Output "`n Try to turn off clutter for  " $account.UserPrincipalName "`n"
			Set-Clutter `
				-Identity $account.UserPrincipalName `
				-Enable $false
			Write-Output ( "Clutter disabled for " + $account.UserPrincipalName )

			Set-UserMailboxPermissions $account.UserPrincipalName
			Write-Output ( "Admin permissions set for " + $account.UserPrincipalName )
		}
		Catch
		{
			# Delay for EOL to sync in the background
			Start-Countdown -Seconds 120

			TRY{
				Write-Output "`n Try to turn off clutter for  " $account.UserPrincipalName "`n"
				$null =Set-Clutter `
					-Identity $account.UserPrincipalName `
					-Enable $false
				Write-Output ( "Clutter disabled for " + $account.UserPrincipalName )

				Set-UserMailboxPermissions $account.UserPrincipalName
				Write-Output ( "Admin permissions set for " + $account.UserPrincipalName )
			}
			Catch
			{
				Write-Output  "----------------- Set-Clutter Exception --------------------"`n
				$_.Exception.Message
				Write-Output "`n  -----------------------------------------------`n"
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
			}
		}
	}
	Remove-PSSession $O365Session
}

# Send email to user with account information
if ($queueMail) {
	Write-Output "`n  ---------------------------- Email Users -----------------------------`n"
	foreach ($mail in $queueMail){
		TRY{
		Write-Output "`n Trying to mail " $mail[0]
		Send-MailMessage `
			-From "{Your support email}" `
			-To $($mail[0]) `
			-Subject "{Organization Name} Account" `
			-BodyAsHtml $($mail[1]) `
			-SmtpServer "{smtp.your.org}"
		}
		CATCH [system.exception]
		{
			Write-Output "`n  ---------------- Email Users $($mail[0]) Exception --------------------`n"
			$_.Exception.Message
			Write-Output "`n  -----------------------------------------------`n"
			$failed += ,@("Attempting to email", $($mail[0]), $($mail[1]), $line.username, $line.id, ($_.Exception.Message + $Myinvocation.ScriptlineNumber))
		}
	}
	Write-Output "`n ------------------------------------------------------------------------`n"
}

# Send email to account creators with summary if anything was processed
if ($sendAdminMail) {

	Write-Output "`n Send mail to admins."

	# Rotate logs older than 6 months
	$log = LogRotate $logPath 6 0 0 0 Append
    $adminBody = "The following accounts were submitted to our account management system:<br />For a log of this transaction look here: <font color='blue'>$($logPath)</font><br /><br />"
    $adminBody += "<html><body><table border=1><tr>"
    $adminBody += "<th>Account Action</th><th>Account type</th><th>Full Name</th><th>User Name</th><th>ID Number</th><th>Notes</th></tr>"

    if ($success.count -gt 0) 	{ $adminBody += "<tr>" }
    foreach ($item in $success) { $adminBody += "<tr><td> $($item[0]) </td> <td>$($item[1])</td><td>$($item[2])</td><td>$($item[3])</td><td>$($item[4])</td><td><font color='green'>$($item[5])</font></td></tr> " }
    if ($success.count -gt 0) 	{ $adminBody += "</tr>" }
    if ($failed.count -gt 0) 		{ $adminBody += "<tr>" }
    foreach ($item in $failed) 	{ $priority = 'High'; $adminBody += "<tr style='background-color: #D7978D;'><td> $($item[0]) </td> <td>$($item[1])</td><td>$($item[2])</td><td>$($item[3])</td><td>$($item[4])</td><td><b>$($item[5])</b></td></tr> " }
    if ($failed.count -gt 0) 		{ $adminBody += "</tr>" }
    $adminBody += "</table>"

    $adminBody += " <br /><br />The following logs were moved to the archive folder: <br /><font color='blue'>$($logPath)\Archive</font><br /><br />"
    $adminBody += "<table border=1>"
    $adminBody += "<tr><th>Log Rotated to the archive folder</th></tr>"

    if ($log.count -gt 0) 		{ $adminBody += "<tr>" }
    foreach ($item in $log) 	{ $priority = 'High'; $adminBody += "<tr style='background-color: #D7978D;'><td><font color='red'>$($item[0])</font></td></tr> " }
    if ($log.count -gt 0) 		{ $adminBody += "</tr>" }

    $adminBody += "</table></body></html>"

	TRY
	{
		Send-MailMessage `
		-From "noreply-AD-AMS@{your org}" `
		-To "accountcreators@{your org}" `
		-Subject "Account Management Summary" `
		-BodyAsHtml $adminBody `
		-SmtpServer "{smtp.your.org}"
		-Priority $priority
	}
	CATCH [system.exception]
	{
		Write-Output "`n  ---------------- Exception : line $(Get-CurrentLine) --------------------`n"
		$_.Exception.Message
		Write-Output "`n  -----------------------------------------------`n"

	}

}
Stop-Logging