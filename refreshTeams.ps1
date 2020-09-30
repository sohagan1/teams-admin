<#

Wrapper script to do whole refresh process

#>

# cd to working directory
cd "" # Path removed

# Connect to VPN so that fetchTTdata can access datastore
c:\windows\system32\rasdial.exe UoE 'sohagan' '' # credentials removed: args are: 1: local VPN connection name, 2: username, 3: VPN password

.\fetchTTdata.ps1

# Disconnect from VPN
c:\windows\system32\rasdial.exe UoE /disconnect

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

.\makeGroupsList.ps1

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

.\guessTeams.ps1

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

.\ingestTimetable.ps1

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

#.\removeAddMembersAll.ps1

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

#.\makeCourseGroupAllocations.ps1

.\checkup.ps1

Read-Host -Prompt "Press any key to continue or CTRL+C to quit"

.\cleanup.ps1