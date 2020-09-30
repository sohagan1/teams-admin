<#
Usage 

Open PowerShell
Run Connect-MicrosoftTeams and login to Teams
Run .\setupTeams.ps1 groups-spec.csv

where groups-spec.csv is a file like this:

Team,OneNote,Channels,Tutors
2020-ILA-Mon-14:10,Y,"M2, LT2, MathsBase","sohagan,tutor2"
2020-ILA-Mon-15:10,N,"M2, LT2, MathsBase","sohagan,tutor2"

#>
$csvpath=$args[0] # this should be a path to a file specifying the groups

# === Read in the data ===
$Header = 'Team', 'OneNote', 'Channels', 'Tutors'
$groupsData = @(Import-Csv -Path $csvpath -Header $Header | select -skip 1) # The @ at the front forces this to be an array even if there is only one row in the csv - makes loops work, select -skip 1 ignores the header row in the file


$groupsData | Out-Host


$decision = $Host.UI.PromptForChoice('Create Teams', 'Are you sure you want to proceed?', @('&Yes'; '&No'), 1)
if ($decision -eq 0) {
    

$nowtime = Get-Date -Format "yyyy-MM-dd-HH-mm"

for ($i=0; $i -lt $groupsData.Count; ++$i) {
		
	# Get the Team name
	$teamName = $groupsData[$i].('Team').Trim() # work out what to call the Team

	#$learnCode = $groupsData[$i].('Group Code') # we'll set the MailNickName to the Learn code to make this Team easy to find later. Note that the MailNickName must be unique so be careful!
		
	# Get an array of channel names to create
	$channels = @($groupsData[$i].('Channels').Split(',').Trim())
		
	# Get an array of owners to add
	$tutors = @($groupsData[$i].('Tutors').Split(',').Trim())

		
	# === Create the Team ===
	# Use the MailNickName property to store the Learn group name so that we can find this Team later
	# -Template "EDU_Class" sets up the Team as a "class" which gives a class notebook
		
	# Decide whether or not to set up a 'Class' Team type
	if ($groupsData[$i].('OneNote').Trim().ToLower() -eq 'y') {
		
		$team = New-Team -DisplayName $teamName -Template "EDU_Class" -AllowCreateUpdateChannels $false -AllowDeleteChannels $false -AllowAddRemoveApps $false -AllowCreateUpdateRemoveTabs $false -AllowCreateUpdateRemoveConnectors $false -ShowInTeamsSearchAndSuggestions $false # not sure if all the additional switches really do anything!
		
	} else {
		
		$team = New-Team -DisplayName $teamName -AllowCreateUpdateChannels $false -AllowDeleteChannels $false -AllowAddRemoveApps $false -AllowCreateUpdateRemoveTabs $false -AllowCreateUpdateRemoveConnectors $false -ShowInTeamsSearchAndSuggestions $false # not sure if all the additional switches really do anything!
		
	}
		
	# Store the new Team's id
	$teamId = $team.GroupId
		
	# Write new Team's ID to a file
	-join ( $teamName, ',', $teamId ) | Out-File '.\team-groupids.csv' -Append
		
	# === Set the team picture === This doesn't work!
	#Set-TeamPicture -GroupId $teamId -ImagePath $avatarPath

	# === Add the channels === 
    if ($channels.Count -gt 0) {		
        Foreach ($channel in $channels) {
		
			New-TeamChannel -GroupId $teamId -DisplayName $channel
		
		}
    }

	# === Add the owners ===
	$filename = -join( '.\', $nowtime, '-', $teamId, '-', $teamName, '-addOwners-failed.txt')
		
    if ($tutors.Count -gt 0) {
        Foreach ($tutor in $tutors) {
			
			# Add @ed.ac.uk to the end of the UUN if it doesn't have it
			if (-not($tutor.contains('@'))) { $tutor = -join ( $tutor, '@ed.ac.uk') }

			try {
				Add-TeamUser -GroupId $teamId -User $tutor -Role 'Owner'
			}
			catch {
				$tutor | Out-File $filename -Append
			}

		}
    }
}


} else {
    Write-Host 'No Teams created.'
}