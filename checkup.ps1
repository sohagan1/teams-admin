param(
    [string] $ClassLists = ''
)
<#

This script takes the output of ingestTimetable.ps1, compares it with team memberships and makes lists of additions and removals.

#>

# Used to store credentials for each functional account used to administer the Teams - not a pretty way of doing it!

$t1p = '' | ConvertTo-SecureString -AsPlainText -Force #password removed
$t2p = '' | ConvertTo-SecureString -AsPlainText -Force #password removed
$t3p = '' | ConvertTo-SecureString -AsPlainText -Force #password removed
$t4p = '' | ConvertTo-SecureString -AsPlainText -Force #password removed
$t5p = '' | ConvertTo-SecureString -AsPlainText -Force #password removed
$t6p = '' | ConvertTo-SecureString -AsPlainText -Force #password removed

$creds = ('mteams1@ed.ac.uk', $t1p),
('mteams2@ed.ac.uk', $t2p),
('mteams3@ed.ac.uk', $t3p),
('mteams4@ed.ac.uk', $t4p),
('mteams5@ed.ac.uk', $t5p),
('mteams6@ed.ac.uk', $t6p) | ForEach-Object {[pscustomobject]@{'User' = $_[0]; 'Pass' = $_[1]}}

function RemoveAdd-MembersByAdmin ($csvpath) {
   
    $errfilename = $csvpath -replace '.csv', '-failed.csv'

    $members = Import-Csv -Path $csvpath

    $removals = @($members | ? 'SideIndicator' -eq '<=')
    $additions = @($members | ? 'SideIndicator' -eq '=>')


	Foreach ($removal in $removals) {
		
		# Change this to suit the input file
		$groupId = $removal.('GroupId')
		$uun  = $removal.('User')
		
		# Add @ed.ac.uk to the end of the UUN if it doesn't have it
		if (-not($uun.contains('@'))) { $uun = -join ( $uun, '@ed.ac.uk') }

		try {
			Remove-TeamUser -GroupId $groupId -User $uun
            "Removed $uun"
		}
		catch {
			$removal | Out-File $errfilename -Append
		}

	}
   
	Foreach ($addition in $additions) {
		
		# Change this to suit the input file
		$groupId = $addition.('GroupId')
		$uun  = $addition.('User')
		
		# Add @ed.ac.uk to the end of the UUN if it doesn't have it
		if (-not($uun.contains('@'))) { $uun = -join ( $uun, '@ed.ac.uk') }

		try {
			Add-TeamUser -GroupId $groupId -User $uun
            "Added $uun"
		}
		catch {
			$addition | Out-File $errfilename -Append
		}

	}

}

if ($ClassLists -eq '') {

    $ClassLists = Get-ChildItem *addMembersToTeams-all.csv | sort LastWriteTime -Descending | select -First 1 | Select -ExpandProperty Name
    $ClassLists = $ClassLists.Insert(0,'.\')

}


for ($i=1; $i -lt 7; ++$i) {
           
    $cred = New-Object System.Management.Automation.PsCredential($creds[$i-1].User,$creds[$i-1].Pass)

    Connect-MicrosoftTeams -Credential $cred

    Start-Sleep -Seconds 1


    $teams = Get-Team -User $creds[$i-1].User
    $csv = import-csv $ClassLists

    $fn = -join (".\", $creds[$i-1].User, "-updates.csv")
    $changes = @()

    $teamsList = import-csv .\team-groupids.csv -Header "DisplayName", "GroupId"
    $teamsToRefresh = $teams | ?{$teamsList.GroupId -Contains $_.'GroupId'}


    foreach ($team in $teamsToRefresh) {

        $members = @(Get-TeamUser -GroupId $team.GroupId -Role Member | ? -Property User -match "s[0-9]{7}@ed.ac.uk" | Select -Property User | Sort User)
        $classList = @($csv | ? -Property GroupId -EQ $team.GroupId | Select -Property User -Unique | Sort User)

        $diff = Compare-Object $members $classList -Property User | Select @{n='GroupId';e={$team.GroupId}}, @{n='DisplayName';e={$team.DisplayName}}, *

        $changes += $diff

        "$($team.DisplayName) has $($diff.count) differences" | Out-Host
        $diff | Out-Host        

    }

    $changes | Export-Csv -Path $fn -Force -NoTypeInformation

    Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

    RemoveAdd-MembersByAdmin($fn)

    Disconnect-MicrosoftTeams

}