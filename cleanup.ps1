
Get-ChildItem *-addMembersToTeams-all.csv | sort LastWriteTime -Descending | Select-Object -skip 1 | Move-Item -Destination ".\archive"
Get-ChildItem *changes.csv | Move-Item -Destination ".\archive"
Get-ChildItem TTDATA*.csv | Move-Item -Destination ".\archive"
