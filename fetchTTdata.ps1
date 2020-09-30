param(
    [string] $filename = ''
)

# Connect to a remote datastore location where BI Suite reports live and copy the latest one to the current directory

$pass = '' | ConvertTo-SecureString -AsPlainText -Force # Password removed
$cred = New-Object System.Management.Automation.PsCredential('ED\',$pass) # Username removed
$path = '' # Path removed

New-PSDrive -name j -Root $path -Credential $cred -PSProvider filesystem

$latest = Get-ChildItem j:\*.csv | sort LastWriteTime -Descending | select -First 1

$source = -join ('j:\', $latest.Name)
#$currDir = Get-Location

Copy-Item -Path $source -Destination '.\' -Force

Remove-PSDrive -Name j