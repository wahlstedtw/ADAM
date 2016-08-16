# AD Account Automation Tool
# =====================================================================
# Much thanks to Nolan Davidson and Jason Frazier for
# all the work you put into the inital script.
# Without your work it would not have grown into the
# tool that it is today.
#
# Authors: Nolan Davidson, Walter Wahlstedt, Jason Frazier
#
# v1.0|06/09/2011:	Account creation script.
# v1.1|09/26/2011:	Added emailing the students their information
# v2.0|05/06/2013:	Added Exchange mailbox creation
#					Implemented new functions: email, sync, SetSecurity
# v2.1|06/12/2013:	Implemented new functions: update, notify, disable,
#					delete, Strip
# v2.2|07/25/2013:	Fixes the Update Function.
# v2.3|05/23/2016:	All emails are sent at the end of the script.
#					Added Set-Logfile function for transcripts.
#					Added date to log file
#					Log file now appends if there are two logs with the same name
#					Delete now only deletes users from pending removals
#					Added Alumni action
#					relocated email and security files
#					Added chars to password generation to force
#					compliance with ad password policy
# v2.4|06/15/2016:	Upgraded to AD Connect.
#					all errors are terminating for try catch
#					staff, faculty and adjunct are now created in office365
#					create fac or staff drives
#					turn clutter off
#
# =====================================================================

# Make all errors terminating
$ErrorActionPreference = Stop

<# ====================================================================
Synopsis:
	Create
		Will create the user account, set mailbox properties, set random password, set distrobution group, OU and Custom attributes for PWM. For 365 users it will assign them a license as well.
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

function Start-Countdown{
<#
	.Synopsis
	 Initiates a countdown on your session.  Can be used instead of Start-Sleep.
	 Use case is to provide visual countdown progress during "sleep" times

	.Example
	 Start-Countdown -Seconds 10

	 This method will clear the screen and display descending seconds

	.Example
	 Start-Countdown -Seconds 10 -ProgressBar

	 This method will display a progress bar on screen without clearing.

	.Link
	 http://www.vtesseract.com/
	.Description
====================================================================
Author(s):		Josh Atwell <josh.c.atwell@gmail.com>
File: 			Start-Countdown.ps1
Date:			2012-04-19
Revision: 		1.0
References:		www.vtesseract.com

====================================================================
Disclaimer: This script is written as best effort and provides no
warranty expressed or implied. Please contact the author(s) if you
have questions about this script before running or modifying
====================================================================
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

Function random-password ( $length = 25 ){
<# ====================================================================

	Generate a random password
	Usage: random-password <length>

========================================================================
#>

        $punc = 46..46
        $digits = 48..57
        $letters = 65..90 + 97..122
        $required = "F4u!"

        # Thanks to
        # https://blogs.technet.com/b/heyscriptingguy/archive/2012/01/07/use-pow
        $password = get-random -count $length `
                -input ($punc + $digits + $letters) |
                        % -begin { $aa = $null } `
                        -process {$aa += [char]$_} `
                        -end {$aa}
        $password = $password + $required
        return $password
}

Function removeGroups ( $username ){
<# ====================================================================

Removes user from all groups and stores those groups in the notes property of the user

========================================================================
#>
	write-host "Remove user groups for " $username

	$DisableUser = Get-QADUser $username
	$groupmemberof = $DisableUser.memberof | Get-QADGroup # Check Group membership and populate the list to Notes

	Foreach ($Group in $groupmemberof){
		$DisNotes = (get-qaduser $username).notes
		# Use $null to suppress unnecessary info in log
		write-host supress
		$null = Set-qaduser $username -notes "$DisNotes $Group;"
	}

	# Remove all memberships from the user except for "Domain Users"
	$DisableUser.memberOf | Get-QADGroup | where {$_.name -notmatch '^users|domain users$'} | Remove-QADGroupMember -member $username
	write-host "Done"
}

function Set-Logfile{
<# ====================================================================

Sets the log file name with the current date

========================================================================
#>
   $date = Get-date -Format 'MM-dd-yyyy'
   return "{path to log file}\Account_Management_Prod_$($date).log"
}

Clear-Host
Add-PSSnapin Quest.ActiveRoles.ADManagement
set-executionpolicy remotesigned

Start-Transcript (set-Logfile) -Append
# Config variables
$priority = 'low'
$parentOU = ''
$year = ''
$baseGroup = ''
$sync = $false
$moveToPendingRemovals = '{inactive users ou}'
$moveToAlumni = '{Alumni users ou}'

# Import csv
$csv = Import-Csv ("{path the account management csv}\account_management.csv")
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
$RootOU = 'OU=Authorized Users,DC={your},DC={domain}'

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
			Write-host ----------------------------- Create Action -----------------------------`n

			if ($line.account_type -eq "none") {
				#$sendUserMail = $false###
				$sendAdminMail = $false
				break
			}
			# Make sure the account type code is valid
			if (($line.account_type -ne "staff") -and ($line.account_type -ne "faculty") -and ($line.account_type -ne "grad") -and ($line.account_type -ne "undergrad") -and ($line.account_type -ne "adjunct") -and ($line.account_type -ne "physicalplant") -and ($line.account_type -ne "pioneerfood")) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Invalid Code: " + $line.account_type))
				Write-host Invalid Code
			}
			# Make sure username isn't already in use.
			elseif ($user = Get-QADUser $line.username) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Duplicate Username: " + $user.samAccountName))
				Write-host "Username in use"
			}
			# Make sure username isn't already in use.
			elseif ($user = Get-QADUser $line.display_name) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Duplicate Display Name: " + $line.display_name))
				Write-host "Display Name in use"
			}
			# Make sure the student/employee ID doesn't already have an account.
			elseif ($user = Get-QADUser -ObjectAttributes @{extensionAttribute1=$line.id}) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Duplicate ID: " + $user.description))
				Write-host "ID already has account"

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
				$passwd = random-password
				$SecurePassword = $passwd | ConvertTo-SecureString -AsPlainText -Force

				if($exchangeCnt){
				Write-host "Connect to exchange for create"
					$username = "{domain\username}"
					$password = cat "{path to credentials}\cred.txt" | convertto-securestring
					$exchangeConnection = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password
					$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange/PowerShell/ -Authentication Kerberos -Credential $exchangeConnection
					Import-PSSession $exchangeSession -DisableNameChecking
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
							-homeDirectory ("\\uc-file\staff$\" + $line.username)`
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
					Write-host "Create User Office 365"
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
						write-host  ----------------- Exception : Line 345 --------------------`n
						$_.Exception.Message
						write-host  ------------------------------------------------`n
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					TRY{
					Set-RemoteMailbox `
						-Identity $line.email `
						-CustomAttribute1 $line.id `
						-CustomAttribute2 $line.ssn_last_4 `
						-CustomAttribute3 $line.birth_date `
						-CustomAttribute5 ($line.id + 'tag')
					}Catch{
						write-host  ----------------- Exception : Line 362 --------------------`n
						$_.Exception.Message
						write-host  ------------------------------------------------`n
						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					TRY{
					Add-DistributionGroupMember `
							-Identity $userCode `
							-Member $line.email `
							-BypassSecurityGroupManagerCheck
					}Catch{
						write-host  ----------------- Exception : Line 375 --------------------`n
						$_.Exception.Message
						write-host  ------------------------------------------------`n

						$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
					}
					if ( ($code -eq "staff") -or ($code -eq "faculty") ) {
					    TRY{
						Set-aduser `
							-Identity $line.username `
							-homeDirectory ("{path to home directory}" + $line.username)`
							-homeDrive Z:
						}Catch{
							write-host  ----------------- set-aduser Exception --------------------`n
							$_.Exception.Message
							write-host  ------------------------------------------------`n
							$failed += ,@($line.account_action, $line.account_type, $line.username, $line.id, $_.InvocationInfo.ScriptLineNumber, ($_.Exception.Message))
						}
					}
					$O365accounts += ,@($code, $line.email)
					$location = "Office 365"
					$sync = $true
				}

				Write-host "Account " $line.username " created"

				$emailFile = "{path to email template}\EmailCreatedUsers.htm"
				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, 'Successfully Created' + $location)
				$sendUserMail = $true
				$emailToUC = $false
			}
			Write-host  -------------------------------------------------------------------------`n
		}
		UPDATE{
			Write-host ----------------------------- Update Action ------------------------------`n
			$code = $line.account_type
			if ($exchangeCnt){
			# adminexchange.txt
				Write-host "Connect to exchange for update"
				$username = "{domain\username}"
				$password = cat "{path to credentials}\cred.txt" | convertto-securestring
				$exchangeConnection = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password
				$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange/PowerShell/ -Authentication Kerberos -Credential $exchangeConnection
				Import-PSSession $exchangeSession
				$exchangeCnt = $false
			}

			Write-host "Attempt to change the account from " $line.display_name " to " $line.new_display_name
			if (Get-Mailbox $line.email){
				if ((Get-QADUser $line.username) -and ($line.new_first_name) -and ($line.new_last_name) -and ($line.new_display_name) -and ($line.new_email) -and ($line.new_username)) {
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

					Rename-Item -path ("{path to home directory}" + $line.username) -newName $line.new_username

					Write-host done
					$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("User changed from " + $line.username + " to " + $line.new_username) )
					$sendUserMail = $true
					$emailToUC = $true
					$emailFile = "{path to email template}\EmailUpdateUsers.htm"
				}
				else{
					write-host "New property is blank"
					$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("New user property is blank"))
				}
			}elseif (Get-RemoteMailbox $line.email){
					Write-host "Update User Office 365"
					if( ($code -eq "grad") ){
						$userCode = "students_grad"
					}
					elseif( ($code -eq "undergrad") ){
						$userCode = "students_undergrad"
					}
					else{ $userCode = $code }

					$rutAddr = $line.username + "@{0365 domain}.mail.onmicrosoft.com"
					Set-User `
						-Identity $line.email `
						-Firstname $line.new_first_name `
						-Lastname $line.new_last_name `
						-initials $line.new_middle_initial `
						-name $line.new_display_name

					Set-RemoteMailbox `
						-Identity $line.email `
						-EmailAddressPolicyEnabled $false `
						-PrimarySMTPAddress $line.new_email `
						-DisplayName $line.new_display_name `
						-UserPrincipalName $line.new_email `
						-samAccountName $line.new_username `
						-alias $line.new_username

					Set-RemoteMailbox -Identity $line.new_email -EmailAddressPolicyEnabled $true

					Add-DistributionGroupMember `
							-Identity $userCode `
							-Member $line.new_email `
							-BypassSecurityGroupManagerCheck

					#$O365accounts += ,@($code, $line.email)
					$location = "Office 365"
					$sync = $true
					Write-host done
					$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("User changed from " + $line.username + " to " + $line.new_username) )
					$sendUserMail = $true
					$emailToUC = $true
					$emailFile = "{path the email template}\EmailUpdateUsers.htm"
			}else{
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Unknown Issue"))
			}
			Write-host  -------------------------------------------------------------------------`n
		}
		NOTIFY {
			Write-host  ----------------------------- Notify Action ------------------------------`n

			# Make sure username exists

			if (!($user = Get-QADUser $line.username)) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $user.samAccountName) )

				Write-host "Account " $line.username " doesn't exist"
			}
			else{
				removeGroups $line.username

				$emailFile = "{path the email template}\EmailNotifyUsers.htm"

				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Removed permissions and notified " + $line.personal_email + ". Disable Pending") )
				write-host ("Notify Account: " + $line.username)
				$sendUserMail = $true
				$emailToUC = $true
			}
			Write-host  ---------------------------------------------------------------------------`n
		 }
		DISABLE {
			Write-host  ----------------------------- Disable Action ------------------------------`n
			$sam = $line.username
			if (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}')) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $line.samAccountName))
				Write-host "Account " $line.username " doesn't exist"
				$sendUserMail = $false
			}
			elseif (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'OU=Pending Removals,DC={domain},DC={name}')){
				removeGroups $line.username

				Try{
				# Disable User account
				# Use $null to suppress unnecessary info in log
				$null = disable-QADuser $line.username

				}CATCH [system.exception]
				{
					write-host  ----------------- Exception --------------------`n
					$_.Exception.Message
					write-host  ------------------------------------------------`n
				   $failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
				}

				# move this user
				TRY{
				write-host "Move User - " $line.username
				$null = Move-QADObject $line.username -NewParentContainer $moveToPendingRemovals
				write-host "Moved"
				}CATCH [system.exception]
				{
					write-host  ----------------- Exception --------------------`n
					$_.Exception.Message
					write-host  ------------------------------------------------`n
				   $failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
				}


				$emailFile = "{path the email template}\EmailDisableUsers.htm"

				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, 'Account disabled and moved to pending removals. Deletion pending.')
				Write-host "Account Disabled"
				$sendUserMail = $false
				$emailToUC = $false
			}
			else{
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account has already been disabled and moved to pending removals"))
				Write-host $line.username " Account has already been disabled and moved to pending removals"
				}
			Write-host  -------------------------------------------------------------------------`n
		 }
		DELETE {
			Write-host  ----------------------------- Delete Action ------------------------------`n
			#$sendUserMail = $false###
			$sam = $line.username

			# Deletes users only from Pending Removals
			if ($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'OU=Pending Removals,DC={domain},DC={name}'){

				TRY{
				write-host "attempting to delete " $sam
				Remove-ADObject $user -Recursive -Confirm:$false
				Write-host "Account " $line.username " Deleted"
				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, 'Account Deleted.')

					if (($line.account_type -eq "undergrad") -or ($line.account_type -eq "grad")){
						$sync = $true
					}
				}CATCH [system.exception]
				{
					write-host  ----------------- Exception --------------------`n
					$_.Exception.Message
					write-host  ------------------------------------------------`n
					$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
				}

			}
			elseif ($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}') {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($user.samAccountName + " is not in pending removals"))

				Write-host "Account is not in pending removals"
			}
			else{
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $line.username))
				Write-host "Account " $line.username " doesn't exist"
			}
			Write-host  -------------------------------------------------------------------------`n
		 }
		SYNC{
			Write-host  ----------------------------- Sync Action ------------------------------`n
			 $sync = $true
			 $success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Syncing AD to cloud") )
			Write-host  -------------------------------------------------------------------------`n
		}
		EMAIL{
			write-host  ----------------------------- Send Test Mail Action ------------------------------`n
			$sendUserMail = $true
			$emailFile = "{path the email template}\EmailCreateduers.htm"
			$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Send Test Email: " + $line.personal_email) )
			Write-host  -------------------------------------------------------------------------`n
		}
		SETSECURITY{
			Write-host Set Security
			$AMS_Ver = read-host -prompt "Account Management Version: "
		 	read-host -prompt "Local password"  -assecurestring |convertfrom-securestring | out-file ("{path the secure files}" + $AMS_Ver + "\local-cred.txt")
			read-host -prompt "Cloud password"  -assecurestring |convertfrom-securestring | out-file ("{path the secure files}" + $AMS_Ver + "\cloud-cred.txt")
		}
		STRIP{
			Write-host  ----------------------------- Strip Action ------------------------------`n
			$sam = $line.username
			if ($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}'){
				removeGroups $line.username
				$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Stripped Permissions - " + $line.username) )
			}
			else{
				Write-host "Account " $line.username " doesn't exist"
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist"))
				}
			Write-host  -------------------------------------------------------------------------`n
		}
		ALUMNI{
			Write-host  ----------------------------- Alumni Action ------------------------------`n
			$sam = $line.username
			$sendUserMail = $false
			# Check to see if user account exists
			if (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'DC={domain},DC={name}')) {
				$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Account doesn't exist: " + $line.samAccountName))
				Write-host "Account " $line.username " doesn't exist"

			}
			# Check to see if account is in pending removals then move to the alumni OU and add them to alumni group
			elseif (!($user = get-aduser -filter {samaccountname -eq $sam} -searchbase 'OU=Pending Removals,DC={domain},DC={name}')){
				# Remove User from groups
				write-host "remove user from groups"
				removeGroups $line.username
				write-host "done"
				# move this user
				write-host "Move User to alumni - " $line.username
				$null = Move-QADObject $line.username -NewParentContainer $moveToAlumni
				write-host 'User Moved'


				if (Get-Mailbox $line.email){
					Add-DistributionGroupMember `
						-Identity alumni `
						-Member $line.email `
						-BypassSecurityGroupManagerCheck

					Write-host done
					$success += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Permissions Removed and moved to Authorized Users/Alumni Users") )

				}
				else{
					write-host "New property is blank"
					$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ("Can't find mailbox"))
				}
			}
			Write-host  -------------------------------------------------------------------------`n
		}
		default {
			if (($line.account_action -eq "none") -or ($line.account_action -eq $null)) {
				write-host  ----------------------------- No Action ------------------------------`n
				$sendAdminMail = $false
				Write-host  -------------------------------------------------------------------------`n
				break
			}
			else{
				write-host ("NO match for the --" + $line.account_action + "--action.")
				$sendAdminMail = $true
			}
		}
	}

	if ($sendUserMail) {

		$userBody = [string]::join([environment]::newline, (Get-Content -path $emailFile))
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
		Write-host "Queue e-mail to " $userTo
		$queueMail += , @($userTo,$userBody)
	}
}

if(!$exchangeCnt){Remove-PSSession $exchangeSession}

if ($sync){
	write-host "sync"
	# Provide username and securestring to server for Remote powershell
	write-host "Waiting to run directory sync"

	Start-Countdown -Seconds 30


	# Creds for connection to server
	Write-host "Connect to server for sync"
	$username = "{domain\username}"
	$password = cat "{path the secure files}Prod\local-cred.txt" | convertto-securestring
	$ADSync = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password

	# Start ADSync on server
	TRY
	{
		write-host "Try to Run Directory Sync on server"
		Invoke-Command -ComputerName "server" -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } -Credential $ADSync
	}
	CATCH [system.exception]
	{
		write-host  ----------------- Exception  : Line 731 --------------------`n
		$_.Exception.Message
		write-host  ------------------------------------------------`n

	   $failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message + ": Line 731"))
	}


	# Delay for ADSync to finish running
	Start-Countdown -Seconds 120
}

if ($O365accounts){
	write-host "O365accounts"
	# read-host -prompt "O365 password"  -assecurestring |convertfrom-securestring | out-file "{path the secure files}Prod\cloud-cred.txt"
	#Connect to Office 365 Remote Powershell
	$pass = cat "{path the secure files}Prod\cloud-cred.txt" | convertto-securestring
	$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist "{o365 admin email}",$pass

	# $appArgs = New-PSSessionOption -ApplicationArguments $O365accounts -SessionOption $appArgs
	$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/  -Authentication Basic -AllowRedirection -Credential $cred
	TRY
	{
		write-host "Try to import PSsession, module and connect to MsolService for O365"
		Import-PSSession $O365Session -DisableNameChecking | out-null
		Import-Module MSOnline
		Connect-MsolService -Credential $cred
	}
	CATCH [system.exception]
	{
		write-host  ----------------- Exception  : Line 759 --------------------`n
		$_.Exception.Message
		write-host  ------------------------------------------------`n

	   $failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message + ": Line 759"))
	}

	write-host "Assign License"
	#Loop through csv to assign user license
	foreach ($item in $O365accounts){

		$code = $($item[0])
		$userPrincipalName = $($item[1])
		write-host $code
		if ($code -eq "none") {
			write-host "A Code of none was found"
			break
		}
		if (($code -eq "grad") -or ($code -eq "undergrad")) {
			$license = '{0365 License Code}'
			write-host "A " $code " license " $license " will be assigned to " $userPrincipalName
		}
		elseif ($code) {
			$license = '{0365 License Code}'
			write-host "A " $code " license " $license " will be assigned to " $userPrincipalName
		}
		TRY
		{
			write-host "Try to set license O365"
			Set-Msoluser -UserPrincipalName $userPrincipalName -UsageLocation US
			Set-MSOLUserLicense -UserPrincipalName $userPrincipalName -AddLicenses $license
			write-host ------test mapi connectivity -----------`n
			Test-MapiConnectivity $userPrincipalName
			write-host -------                     ------------`n
		}
		CATCH [system.exception]
		{
		    write-host  ----------------- Exception  : Line 795  --------------------`n
			$_.Exception.Message
			write-host  ------------------------------------------------`n

		   $failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message + ": Line 795" ))
		}
		write-host ( $code + " license assigned to user " + $userPrincipalName)
		TRY{
			write-host "turn off clutter"
			Set-Clutter `
				-Identity $line.email `
				-Enable $false
		}Catch{
			write-host  "----------------- Set-Clutter Exception --------------------"`n
			$_.Exception.Message
			write-host  ------------------------------------------------`n
			$failed += ,@($line.account_action, $line.account_type, $line.display_name, $line.username, $line.id, ($_.Exception.Message))
		}
		write-host ( "Clutter disabled for" + $userPrincipalName )
	}
	Remove-PSSession $O365Session
}

# Send email to user with account information
if ($queueMail) {
	write-host  ----------------------------- Email Users -----------------------------`n
	foreach ($mail in $queueMail){
		TRY{
		write-host "Trying to mail " $mail[0]
		Send-MailMessage `
			-From "{your from email}" `
			-To $($mail[0]) `
			-Subject "{Organization Name} Account" `
			-BodyAsHtml $($mail[1]) `
			-SmtpServer "{mail.yoursmpt.ext}"

		}
		CATCH [system.exception]
		{
			write-host  ----------------- Email Users $($mail[0]) Exception --------------------`n
			$_.Exception.Message
			write-host  ------------------------------------------------`n
			$failed += ,@("Attempting to email", $($mail[0]), $($mail[1]), $line.username, $line.id, ($_.Exception.Message + ": Line 822"))
		}
	}
	Write-host  -------------------------------------------------------------------------`n
}

# Send email to account creators with summary if anything was processed, mark as high priority if there was a failure
if ($sendAdminMail) {

	write-host Send mail to admins

    $adminBody = "The following accounts were submitted to our account management system:`r`r"
    $adminBody += "<html><body><table border=1><tr>"
    $adminBody += "<th>Account Action</th><th>Account type</th><th>Full Name</th><th>User Name</th><th>ID Number</th><th>Notes</th></tr>"
write-host $adminBody
    if ($success.count > 0) 	{ $adminBody += "<tr>" }
    foreach ($item in $success) { $adminBody += "<tr><td> $($item[0]) </td> <td>$($item[1])</td><td>$($item[2])</td><td>$($item[3])</td><td>$($item[4])</td><td><font color='green'>$($item[5])</font></td></tr> " }
    if ($success.count > 0) 	{ $adminBody += "</tr>" }
    if ($failed.count > 0) 		{ $adminBody += "<tr>" }
    foreach ($item in $failed) 	{ $priority = 'High'; $adminBody += "<tr style='background-color: #D7978D;'><td> $($item[0]) </td> <td>$($item[1])</td><td>$($item[2])</td><td>$($item[3])</td><td>$($item[4])</td><td><font color='red'>$($item[5])</font></td></tr> " }
    if ($failed.count > 0) 		{ $adminBody += "</tr>" }

    $adminBody += "</table></body></html>"
write-host $priority
	TRY
	{
		Send-MailMessage `
		-From "{your from email}" `
		-To "{your admin email group}" `
		-Subject "Account Management Summary" `
		-BodyAsHtml $adminBody `
		-SmtpServer "{mail.yoursmpt.ext}" `
		-Priority $priority
	}
	CATCH [system.exception]
	{
		write-host  ----------------- Exception : Line 861 --------------------`n
		$_.Exception.Message
		write-host  ------------------------------------------------`n

	}

}
stop-transcript