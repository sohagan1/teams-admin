param(
    [string] $csvpath
)
<#

This script takes a csv file and does member additions and removals.

The csv should have these columns:

GroupId, DisplayName, User, SideIndicator
9a53eeae-2073-46c6-8791-144da6061860, Team Name, uun@ed.ac.uk, <=

SideIndicator is the output of a Compare-Object commandlet.

<= indicates that the user is a member of the Team but not on the class list
=> indicates that the user is not a member of the Team but is on the class list

#>

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
RemoveAdd-MembersByAdmin($csvpath)