#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################


# Let’s start by getting all of the events in this event channel, then categorize them

$Monitoring = Get-WinEvent -ComputerName <server name> -LogName "Microsoft-Exchange-ManagedAvailability/Monitoring"
$ErrorEvents = $Monitoring | Where-Object {$_.LevelDisplayName -eq "Error" -and $_.TimeCreated -gt [datetime]::Now.AddDays(-2)}
$HealthyEvents = $Monitoring | Where-Object {$_.LevelDisplayName -eq "Information" -and $_.TimeCreated -gt [datetime]::Now.AddDays(-2)}

# For each error event, let’s see if there is a newer healthy event for the same health set

$UnhealthyHealthSets = @()
foreach ($ErrorEvent in $ErrorEvents)
{
	# We'll skip converting this into XML and instead reference by index for better performance
	$HealthSet = $ErrorEvent.Properties[0].Value
	if ($UnhealthyHealthSets.Contains($HealthSet))
	{
		# We have already included a more recent failure
		continue
	}
	$StillFailed = $true
	$HealthyEventsAfterError = $HealthyEvents | Where-Object {$_.TimeCreated -gt $Error.TimeCreated}
	foreach ($SuccessEvent in $HealthyEventsAfterError)
	{
		if ($SuccessEvent.Properties[0].Value -eq $ErrorEvent.Properties[0].Value)
		{
			$StillFailed = $False
			break
		}
	}
	if ($StillFailed)
	{
		$UnhealthyHealthSets += $ErrorEvent.Properties[0].Value
		$ErrorEvent.Properties[0].Value
		# Consider outputting $ErrorEvent.Message to a text or HTML file for easy consumption
	}
}
