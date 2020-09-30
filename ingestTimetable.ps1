param(
    [string] $tutgroupscsv = '.\tutorialteams-all.csv'
)

<# 

This script imports the latest data from timetabling and writes out an addMembersToTeams-all.csv
containing information about which students should be in which teams.

#>



function Read-TTData()  {
    
    # Find the latest TTDATA in the current directory
    $latest = Get-ChildItem TTDATA*.csv | sort LastWriteTime -Descending | select -First 1
    $ttfilename = $latest.Name

    $Global:ttdata = Import-Csv -Path $ttfilename | Where {$_.('UUN (upper case)') -like "S*"}

    Write-Host "Read: $ttfilename"

}

function Read-TeamsAdmins($tapath) {

    $Global:teamsadmins = Import-Csv -Path $tapath

}

function Read-TutorialTeams($tutteamspath) {
    
    $Global:tutorial_teams = Import-Csv -Path $tutteamspath |
                                Select-Object *, @{n='TeamShort';e={$_.('Team')  -replace '[^a-zA-Z0-9]', '' }}
    Write-Host "Read: $tutteamspath"

}

function Read-TeamGroupIds($teamgroupidspath) {
    
    $Global:team_groupids = Import-Csv -Path $teamgroupidspath -Header "DisplayName", "GroupId" |
                                Select-Object *, @{n='DNshort';e={$_.('DisplayName')  -replace '[^a-zA-Z0-9]', '' }}
    Write-Host "Read: $teamgroupidspath"

}

function Write-AddMembersToTeams {

    $nowdatetime = Get-Date -Format "yyyy-MM-dd-HH-mm"
    $Global:addMembersFN = -join ('.\', $nowdatetime, '-addMembersToTeams-all.csv')
    New-Item -Path $addMembersFN -ItemType File
    "TeamsAdmin, GroupId, ShortDisplayName, ShortActivityName, User" | Out-File $addMembersFN -Append

    for ($i=0; $i -lt $Global:ttdata.Count; ++$i) {
        
        $uun = -join ($global:ttdata[$i].('UUN (upper case)'), '@ed.ac.uk')
        $ahk = $global:ttdata[$i].'Activity Host Key'
        $courseCode = $global:ttdata[$i].('Course Code').Split('_')[0]
        
        # Look up the Teams admin
        try {
			$getTeamsAdmin = $global:teamsadmins | Where {$_.('Course Code') -eq $courseCode} | Select -Property 'TeamsAccount' | Select -first 1
            #$lineout = @($ahk, $uun)

		} catch {
			#$lineout = @("Error finding Team for this tutorial", $ahk, $uun)
            #$error1 = true
		}

        #
        try {
			$getTeamShort = $global:tutorial_teams | Where {$_.('Host Key') -eq $ahk} | Select -Property 'TeamShort' | Select -first 1
            $lineout = @($ahk, $uun)

		} catch {
			$lineout = @("Error finding Team for this tutorial", $ahk, $uun)
            $error1 = true
		}
        
        #
        try {
			$getGroupId = $global:team_groupids | Where {$_.DNshort -eq $getTeamShort.TeamShort} | Select -Property 'GroupId','DNshort' | Select -first 1
            $lineout = @($getTeamsAdmin[0].TeamsAccount, $getGroupId[0].GroupId, $getGroupId[0].DNshort) + $lineout
		} catch {
            if (-not $error1) {
			    $lineout = "Not found", $global:ttdata[$i]
            }
		}
        
        #"$($i): $lineout" | Out-Host
        $lineout -join ',' | Out-File $addMembersFN -Append

    }

}

if ([string]::IsNullOrWhiteSpace($tutgroupscsv)) { $tutgroupscsv = '.\tutorialteams-all.csv' }


$Global:ttdata = $null
$Global:teamsadmins = $null
$Global:tutorial_teams = $null
$Global:team_groupids = $null
$Global:addMembersFN = $null

Read-TTData
Read-TeamsAdmins('.\course-teamsaccount.csv')
Read-TutorialTeams($tutgroupscsv)
Read-TeamGroupIds('.\team-groupids.csv')
Write-AddMembersToTeams