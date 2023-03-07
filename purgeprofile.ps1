do {
    $Profiles = Get-WmiObject -Class Win32_UserProfile
    $Index = 1
    $ProfilesArray = @{}
    foreach ($Profile in $Profiles) {
        $UserName = $Profile.LocalPath.Split("\")[-1]
        $ProfilesArray[$Index] = $UserName
        Write-Host "$Index. $UserName"
        $Index++
    }

    $Selection = Read-Host "Enter the number of the profile you want to purge"
    if ($ProfilesArray.ContainsKey([int]$Selection)) {
        $User = $ProfilesArray[[int]$Selection]
        $Profile = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.LocalPath -like "C:\Users\$User" }

        if ($Profile) {
            $Size = "{0:N2} MB" -f ((Get-ChildItem $Profile.LocalPath -Recurse -Force | Measure-Object Length -Sum).Sum / 1MB)
            Write-Host "The profile for user $User is taking up $Size of disk space."

            $Confirmation = Read-Host "Do you want to delete the entire profile for $User? (Y/N)"
            if ($Confirmation -eq "Y") {
                Remove-WmiObject -InputObject $Profile -ErrorAction SilentlyContinue
                Write-Host "The profile for user $User has been deleted."
            } else {
                Write-Host "The profile for user $User was not deleted."
            }
        } else {
            Write-Host "No profile was found for user $User."
        }
    } else {
        Write-Host "Invalid selection"
    }

    $Continue = Read-Host "Do you want to purge another profile? (Y/N)"
} while ($Continue -eq "Y")
