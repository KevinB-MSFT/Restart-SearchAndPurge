#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Loops through the search and purge process in eDiscovery to work around the 10 mail item limit
# Restart-SearchAndPurge_1.0.ps1
#  
# Created by: Kevin Bloom 2/5/2021 Kevin.Bloom@Microsoft.com 
#
#Version change log:
#1.0: Initial script creation
#1.1: Added extra logging and corrected logic in the "While ($CompletionStatus -ne "Completed")" loop
#
#########################################################################################
Function Write-Log {
	Param ([string]$string)
	$NonInteractive = 1
	# Get the current date
	[string]$date = Get-Date -Format G
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	# If NonInteractive true then supress host output
	if (!($NonInteractive)){
		( "[" + $date + "] - " + $string) | Write-Host
	}
}

# Sleeps X seconds and displays a progress bar
Function Start-SleepWithProgress {
	Param([int]$sleeptime)
	# Loop Number of seconds you want to sleep
	For ($i=0;$i -le $sleeptime;$i++){
		$timeleft = ($sleeptime - $i);
		# Progress bar showing progress of the sleep
		Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		# Sleep 1 second
		start-sleep 1
	}
	Write-Progress -Completed -Activity "Sleeping"
}

# Setup a new O365 Powershell Session using RobustCloudCommand concepts
Function New-CleanIPPSSession {
	#Prompt for UPN used to login to EXO 
   Write-log ("Removing all PS Sessions")

   # Destroy any outstanding PS Session
   Get-PSSession | Remove-PSSession -Confirm:$false
   
   # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
   [System.GC]::Collect()
   
   # Sleep 10s to allow the sessions to tear down fully
   Write-Log ("Sleeping 10 seconds to clear existing PS sessions")
   Start-Sleep -Seconds 10

   # Clear out all errors
   $Error.Clear()
   
   # Create the session
   Write-Log ("Creating new PS Session")
	
   # Check for an error while creating the session
	If ($Error.Count -gt 0){
		Write-log ("[ERROR] - Error while setting up session")
		Write-log ($Error)
		# Increment our error count so we abort after so many attempts to set up the session
		$ErrorCount++
		# If we have failed to setup the session > 3 times then we need to abort because we are in a failure state
		If ($ErrorCount -gt 3){
			Write-log ("[ERROR] - Failed to setup session after multiple tries")
			Write-log ("[ERROR] - Aborting Script")
			exit		
		}	
		# If we are not aborting then sleep 60s in the hope that the issue is transient
		Write-log ("Sleeping 60s then trying again...standby")
		Start-SleepWithProgress -sleeptime 60
		
		# Attempt to set up the sesion again
		New-CleanIPPSSession
	}
   
   # If the session setup worked then we need to set $errorcount to 0
   else {
	   $ErrorCount = 0
   }
   # Import the PS session/connect to EXO
	$null = Connect-IPPSSession -UserPrincipalName $LogonUPN 
   # Set the Start time for the current session
	Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy; Goes ahead and resets it every "$ResetSeconds" number of seconds (14.5 mins) either way 
Function Test-IPPSSession {
	# Get the time that we are working on this object to use later in testing
	$ObjectTime = Get-Date
	# Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession
	# Make sure we found a session
	if ($SessionInfo -eq $null) { 
		Write-log ("[ERROR] - No Session Found")
		Write-log ("Recreating Session")
		New-CleanIPPSSession
	}	
	# Make sure it is in an opened state if not log and recreate
	elseif ($SessionInfo.State -ne "Opened"){
		Write-log ("[ERROR] - Session not in Open State")
		Write-log ($SessionInfo | fl | Out-String )
		Write-log ("Recreating Session")
		New-CleanIPPSSession
	}
	# If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
		Write-log ("Rebuilding Connection")
		
		# Estimate the throttle delay needed since the last session rebuild
		# Amount of time the session was allowed to run * our activethrottle value
		# Divide by 2 to account for network time, script delays, and a fudge factor
		# Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		[int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		# If the delay is >15s then sleep that amount for throttle to recover
		if ($DelayinSeconds -gt 0){
			Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
			Start-SleepWithProgress -SleepTime $DelayinSeconds
		}
		# If the delay is <15s then the sleep already built into New-CleanIPPSSession should take care of it
		else {
			Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		}
		# new O365 session and reset our object processed count
		New-CleanIPPSSession
	}
	else {
		# If session is active and it hasn't been open too long then do nothing and keep going
	}
	# If we have a manual throttle value then sleep for that many milliseconds
	if ($ManualThrottle -gt 0){
		Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
		Start-Sleep -Milliseconds $ManualThrottle
	}
}

##Start Script##

#Set Variables
$logfilename = '\Restart-SearchAndPurge'
#$outputfilename = '\Restart-SearchAndPurge_Output_'
$execpol = get-executionpolicy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force  #this is just for the session running this script
Write-Host;$LogonUPN=Read-host "Type in UPN for account that will execute this script"
$Tenant=Read-host "Type in your tenant domain name (eg <domain>.onmicrosoft.com)";write-host "...pleasewait...connecting to IPPS..."
# Set $OutputFolder to Current PowerShell Directory
[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))
$outputFolder = [IO.Directory]::GetCurrentDirectory()
$DateTicks = (Get-Date).Ticks
$logFile = $outputFolder + $logfilename + $DateTicks + ".txt"
#$OutputFile= $outputfolder + $outputfilename + $DateTicks + ".csv"
[int]$ManualThrottle=0
[double]$ActiveThrottle=.25
[int]$ResetSeconds=870

$Counter0 = 1 #Number of iterations of script.  So if there are 1000 mails in a mailbox(s) then you need to set to 1000/10 = 100 Etc.
##Ensure you check over content search to make sure that the query used pulls the emails you wish to remove before running this to purge.  Specifically if you run the purge as HardDelete

$Progress0 = 0 #Progres monitor and condition for breaking loop
$SearchName = "TestScriptPurge" #Name of the search created in content search
$PurgeName = $SearchName+"_Purge" #all purges are appended with _purge 
$PurgeSetting = "softDelete" #HardDelete or SoftDelete depending on what you need
$CompletionStatus = $Null #Declaration of condition used to check status 

# Setup our first session to O365
$ErrorCount = 0
New-CleanIPPSSession
Write-Log ("Connected to IPPS Online")
write-host;write-host -ForegroundColor Green "...Connected to IPPS Online as $LogonUPN";write-host

# Get when we started the script for estimating time to completion
$ScriptStartTime = Get-Date
$startDate = Get-Date
write-progress -id 1 -activity "Beginning..." -PercentComplete (1) -Status "initializing variables"

# Clear the error log so that sending errors to file relate only to this run of the script
$error.clear()

Write-Host "Running Compliance Purge $Counter0 times" -ForeGroundColor Red

do { #Do until all iterations of set in $counter0 have completed

	Write-log ("if (!(!(Get-ComplianceSearchAction $PurgeName -erroraction 0)))  #Need to check if a purge action exists for the search name then remove it")
    Test-IPPSSession
    if (!(!(Get-ComplianceSearchAction $PurgeName -erroraction 0))) { #Need to check if a purge action exists for the search name then remove it
    
        #Remove Current Search Action before continuing
        "Found Current Search Action - Removing before Proceeding"
        Write-log ("Remove-ComplianceSearchAction $PurgeName -Confirm:$false")
		Remove-ComplianceSearchAction $PurgeName -Confirm:$false
    }
    
    "Running Purge"
    Write-log ("New-ComplianceSearchAction -Purge -PurgeType $PurgeSetting -SearchName $SearchName -Confirm:$false -Force #Perform Purge")
	Test-IPPSSession
    New-ComplianceSearchAction -Purge -PurgeType $PurgeSetting -SearchName $SearchName -Confirm:$false -Force #Perform Purge
    
    Start-Sleep -Seconds 5

	Test-IPPSSession
    $CompletionStatus = (Get-ComplianceSearchAction $PurgeName).Status #Pull initial status to feed while loop / short circuit if it is completed before loop
    Write-log ("Pre-loop $CompletionStatus = (Get-ComplianceSearchAction $PurgeName).Status")
	"Checking Status until completed"
    While ($CompletionStatus -ne "Completed") { #If purge takes awhile this periodically checks status until completed

        Start-Sleep -Seconds 2
        Test-IPPSSession
        $CompletionStatus = (Get-ComplianceSearchAction $PurgeName).Status
		Write-log ("In-loop $CompletionStatus = (Get-ComplianceSearchAction $PurgeName).Status")
		

    }
    $Progress0++
    $PercentageDone0 = [Math]::Round(($Progress0/$Counter0), 2)*100
    Write-Progress -Activity "Purge in Progress..." -PercentComplete $PercentageDone0 -CurrentOperation "$($Progress0) of $($Counter0); $($PercentageDone0)% complete." -Status "Progress"
	Write-log ("Purge in Progress... Percent Complete $PercentageDone0 CurrentOperation $($Progress0) of $($Counter0); $($PercentageDone0)% complete. Status Progress")

} While ($Progress0 -lt $Counter0) #Ends Do While based on if progress has reached total iterations

Write-log ("Script Completed")
"Script Completed"
    
   
