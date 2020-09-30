param(
    [string] $csvpath = '.\groups-list.csv'
) 

# Load course info to look up shortcodes
$courseInfo = Import-Csv -Path .\course-teamsaccount.csv
"Imported list of courses that use Teams" | Out-Host

# Load teamgroupids to look for Teams
$teamgroupids = Import-Csv -Path .\team-groupids.csv -Header 'DisplayName', 'GroupId' | Select-Object *, @{Name='SmashedDisplayName';Expression={($_.DisplayName -replace '[^a-zA-Z0-9]', '').ToLower()}}
"Imported list of Teams" | Out-Host

$ignore = Import-Csv '.\activities-ignore.csv' | %{$_.Name}
"Imported list of activites to ignore on timetable" | Out-Host

$manualRewrites = Import-Csv '.\activities-manualRewrites.csv'
"Imported list of manual rewrites to timetable activities" | Out-Host

$ttgroups = Import-Csv $csvpath -Header 'Name', 'Module Host Key', 'Host Key', 'Activity Type Name', 'Scheduled Days', 'Scheduled Start Time', 'Scheduled End Time', 'Allocated Location Name' |
              Where {$ignore -notcontains $_.('Name')} | Select-Object *,'Team' -Skip 1
"Importing list of timetable activities" | Out-Host


"Starting to match timetable activities to Teams" | Out-Host
# Try to construct the name of the Team from the data in the row
foreach ($ttgroup in $ttgroups) {

    # Check for manual override

    $foundManRewrite = @($manualRewrites | Where-Object {$_.'Name' -eq $ttgroup.Name -and $_.'Scheduled Days' -eq $ttgroup.'Scheduled Days' -and $_.'Scheduled Start Time' -eq $ttgroup.'Scheduled Start Time'})

    if ($foundManRewrite.Count -eq 1) {

        $ttgroup.Team = $foundManRewrite[0].('Team')

    } else {
    # Otherwise guess


        $courseCode = $ttgroup.('Module Host Key').Split('_')[0]
        $getShortCode = @($courseInfo | Where {$_.('Course Code') -eq $courseCode} | Select -Property 'School Acronym Suffix' | Select -first 1)

        # if no short code then this course is not using Teams, so remove the object grom ttgroup and move on to the next row.
        if ($getShortCode.Count -eq 0) {

            $ttgroup.Team = 'NoTeams'
            continue

        }
        

        $shortCode = $getShortCode.'School Acronym Suffix'
        $ttgroup.Team = ''

        # Make guess with 3-letter days
        if ($ttgroup.('Scheduled Start Time') -eq '9:00') {

            $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,3), '09')

        } else {

            $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,3), $ttgroup.('Scheduled Start Time').Substring(0,2))

        }
        
        #Write-Host $guess

        # Compare against list of smashed Team names
        $found = @($teamgroupids | Where {$_.('SmashedDisplayName').Contains($guess.tolower())} | Select -Property 'DisplayName')

        #$guess | Out-Host
        #$found.Count | Out-Host

        if ($found.Count -eq 1) {

            $ttgroup.Team = $found[0].('DisplayName') 

            #$found | Out-Host 
            #$ttgroup.Team | Out-Host

        } else {
            
            # Make guess with 4-letter days
            if ($ttgroup.('Scheduled Start Time') -eq '9:00') {

                $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,4), '09')

            } else {

                $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,4), $ttgroup.('Scheduled Start Time').Substring(0,2))

            }

            #Write-Host $guess
            $found = @($teamgroupids | Where {$_.('SmashedDisplayName').Contains($guess.tolower())} | Select -Property 'DisplayName')

            if ($found.Count -eq 1) {

                $ttgroup.Team = $found[0].('DisplayName') 

            } else {

                # Make guess with 5-letter days
                if ($ttgroup.('Scheduled Start Time') -eq '9:00') {

                    $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,5), '09')

                } else {

                    $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,5), $ttgroup.('Scheduled Start Time').Substring(0,2))

                }

                $found = @($teamgroupids | Where {$_.('SmashedDisplayName').Contains($guess.tolower())} | Select -Property 'DisplayName')

                if ($found.Count -eq 1) {

                    $ttgroup.Team = $found[0].('DisplayName')

                } else {

                    # Maybe it's a 9:00 start but the time isn't 0 padded in Team name
                    if ($ttgroup.('Scheduled Start Time') -eq '9:00') {
                    
                        # Make guess with 3-letter days
                        $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,3), "9")
                        #Write-Host $guess
                    
                        $found = @($teamgroupids | Where {$_.('SmashedDisplayName').Contains($guess.tolower())} | Select -Property 'DisplayName')
                        #Write-Host $found.Count

                        if ($found.Count -eq 1) {

                            $ttgroup.Team = $found[0].('DisplayName') 
                            #Write-Host $ttgroup.Team

                        } else {

                            # Make guess with 4-letter days
                            $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,4), "9")
                            #Write-Host $guess

                            $found = @($teamgroupids | Where {$_.('SmashedDisplayName').Contains($guess.tolower())} | Select -Property 'DisplayName')
                            #Write-Host $found.Count

                            if ($found.Count -eq 1) {

                                $ttgroup.Team = $found[0].('DisplayName')
                                #Write-Host $ttgroup.Team

                            } else {

                                # Make guess with 5-letter days
                                $guess = -join @('2020', $shortCode, $ttgroup.('Scheduled Days').substring(0,5), "9")
                                #Write-Host $guess

                                $found = @($teamgroupids | Where {$_.('SmashedDisplayName').Contains($guess.tolower())} | Select -Property 'DisplayName')
                                #Write-Output $found.Count

                                if ($found.Count -eq 1) {

                                    $ttgroup.Team = $found[0].('DisplayName')
                                    #Write-Host $ttgroup.Team

                                }

                            }

                        }

                    }

                }

            }

        }

    }

}

$unmatched = @($ttgroups | Where{$_.Team -eq ''}).Count

if ($unmatched) {

    Write-Error "There were $unmatched unmatched activities"

} else {

    "All activities matched to a Team" | Out-Host

}

$ttgroups | Where{$_.Team -notmatch 'NoTeams'} | Export-Csv .\tutorialteams-all.csv -NoTypeInformation