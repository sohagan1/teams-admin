$t1p = '' | ConvertTo-SecureString -AsPlainText -Force  # Pasword removed
$t2p = '' | ConvertTo-SecureString -AsPlainText -Force  # Pasword removed
$t3p = '' | ConvertTo-SecureString -AsPlainText -Force  # Pasword removed
$t4p = '' | ConvertTo-SecureString -AsPlainText -Force  # Pasword removed
$t5p = '' | ConvertTo-SecureString -AsPlainText -Force  # Pasword removed
$t6p = '' | ConvertTo-SecureString -AsPlainText -Force  # Pasword removed

$creds = ('mteams1@ed.ac.uk', $t1p),
('mteams2@ed.ac.uk', $t2p),
('mteams3@ed.ac.uk', $t3p),
('mteams4@ed.ac.uk', $t4p),
('mteams5@ed.ac.uk', $t5p),
('mteams6@ed.ac.uk', $t6p) | ForEach-Object {[pscustomobject]@{'User' = $_[0]; 'Pass' = $_[1]}}


for ($i=1; $i -lt 7; ++$i) {

    $fn = @(Get-ChildItem *mteams$i-changes.csv | Select Name)
    
    if ($fn.Count -eq 1) {
        
        $file = -join ('.\' , $fn[0].Name)

        $file | Out-Host

#       $decision = $Host.UI.PromptForChoice('Make changes', "Do you want to process this file?", @('&Yes'; '&No'), 1)
#       if ($decision -eq 0) {
            
            $cred = New-Object System.Management.Automation.PsCredential($creds[$i-1].User,$creds[$i-1].Pass)

            Connect-MicrosoftTeams -Credential $cred

            Start-Sleep -Seconds 1

            .\removeAddMembersByAdmin.ps1 $file

            Disconnect-MicrosoftTeams


#    } else {



#    }

    } else {

        "Something is wrong: there are more changes files than expected" | Out-Host

    }

}