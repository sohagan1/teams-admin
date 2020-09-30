param(
    [string] $in = '',
    [String] $out = '.\groups-list.csv'
)

if ($in -eq '') {
    
    # look for the latest TTdata file in the current directory

    $latest = Get-ChildItem TTDATA*.csv | sort LastWriteTime -Descending | select -First 1

    $in = $latest.Name

}

$data = Import-Csv $in

$gps = $data | Group-Object 'Activity Host Key' |
        %{ $_.Group | Select 'Activity Name', 'Course Code' ,'Activity Host Key', 'Actvity Type', 'Day Name', 'Start Time', 'End Time', 'Location Name' -First 1 } |
            Sort 'Course Code', 'Activity Name'

$gps | Export-Csv $out -NoTypeInformation