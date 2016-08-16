# AD-account-automation

This tool was created to automate the lifecycle of a Microsoft Active directory account. Written in powershell it has functions such as Create, Update and Delete. It takes a csv as input and is meant to be run as a scheduled task.

Synopsis:
	Create
		Will create the user account, set mailbox properties, set random password, set distribution  group, OU and Custom attributes for PWM. For 365 users it will assign them a license as well.
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