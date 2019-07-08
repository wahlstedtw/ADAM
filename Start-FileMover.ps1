#
# File Mover Tool
# Authors: Walter Wahlstedt
#
# v1.0|12/10/2018:	Initialization
#
# =====================================================================

# Make all errors terminating
$ErrorActionPreference = "Stop";

<# ====================================================================
Synopsis:
  Moves files from a folder on the server to a share
=======================================================================
#>

Clear-Host
# Import AMS general functions and Active Directory modules.
try {
	# Import-Module ActiveDirectory
	Import-Module ADAMFunctions -force
	# Get name and path of the current script.
	$ScriptInfo = Get-ScriptInfo
	$ScriptInfo.Path
	$ScriptInfo.Name
	# Records a log of the script output.
	$logInfo = Start-Logging $ScriptInfo.Path $ScriptInfo.Name 6 0 0 0 0 append
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

	Define variables

========================================================================
#>

$FileShare = "{file share where you want to move the file to}"
$csv = Import-Csv ("$ScriptInfo.Path\$ScriptInfo.Name\DataFiles\file_submissions.csv")


$CSV | ForEach-Object {
  # Get all items (should be only one) in the location and rename them to newName
  TRY{
      Get-ChildItem -Path $_.location | Rename-Item -NewName $_.newName
  }
  CATCH [system.exception]
  {
    Write-Output "`n  ---------------- Rename file failed : $(Get-CurrentLine) --------------------`n"
    $_.Exception.Message
    Write-Output "`n  -----------------------------------------------`n"
  }

  #  Test if the file path on the syllabi share exists and create if not
  if(!(Test-Path -Path "$FileShare\$($_.year)\$($_.Term)")){
    Write-Information "Path Doesn't Exist. Prepare for creation"
    TRY{
      New-Item "$FileShare\$($_.year)\$($_.Term)" -type directory
		}
	  CATCH [system.exception]
		{
			Write-Output "`n  ---------------- Create New share file failed : $(Get-CurrentLine) --------------------`n"
			$_.Exception.Message
			Write-Output "`n  -----------------------------------------------`n"
		}
  }

  # Move all items (should be only one) from location to
  # the share and create the folders if they don't exist
  TRY{
    Get-ChildItem -Path $_.location | Move-Item -Force -Destination "$FileShare\$($_.year)\$($_.Term)"
	}
	CATCH [system.exception]
	{
    Write-Output "`n  ---------------- Move file failed : $(Get-CurrentLine) --------------------`n"
    $_.Exception.Message
    Write-Output "`n  -----------------------------------------------`n"
	}

    # Delete the folder from jics server

    TRY{
        Remove-Item $_.location
       }
        CATCH [system.exception]
        {
        Write-Output "`n  ---------------- Delete folder failed : $(Get-CurrentLine) --------------------`n"
        $_.Exception.Message
        Write-Output "`n  -----------------------------------------------`n"
        }
}

Stop-Logging