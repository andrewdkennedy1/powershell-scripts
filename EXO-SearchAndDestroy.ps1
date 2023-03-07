Import-Module ExchangeOnlineManagement
Connect-IPPSSession



###
$retentiondays = 14 # number of days to keep the results of the Compliance Search - applies to Search, Export, and Purge results. Adjust if needed.
###
$elapsed = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host "When was that message received?"
Write-Host "1. Today."
Write-Host "2. I will specify the date."`n
$choice1 = Read-Host "Please select an option that fits your requirements the best ( 1 or 2 )"
while("1","2" -notcontains $choice1) {
    
    $choice1 = Read-Host "Please select an option that fits your requirements the best ( 1 or 2 )"
}
If ($choice1 -eq '1') {
$startdate = (Get-Date).Date
}
ElseIf ($choice1 -eq '2') {
$startdate = Read-Host `n"Please specify the date. The search will be done from that date"
}
Write-Host `n"Two options to choose from:"
Write-Host "1. I have all information I need for a manual input: the sender and the subject."
Write-Host "2. I know which user got the message of interest: all information will be used automatically once that message is located by me."`n
$choice2 = Read-Host "Please select an option that fits your requirements the best ( 1 or 2 )"
while("1","2" -notcontains $choice2) {
    
    $choice2 = Read-Host "Please select an option that fits your requirements the best ( 1 or 2 )"
}
If ($choice2 -eq '1') {
$sendere = Read-Host `n"Sender's email address"
    $subject = Read-Host "Part of the subject line"
}
ElseIf ($choice2 -eq '2') {
    
    $enddate = get-date
        
    $emailaddress = Read-Host `n"Email of the user that got the message of interest"
$collection = @()
    
    $i = 0
$emails = Get-MessageTrace -StartDate $startdate -EndDate $enddate -RecipientAddress $emailaddress | Select-Object Received,SenderAddress,Subject
ForEach ($email in $emails) {
$outObject = "" | Select-Object Number,Received,"Sender Address",Subject
    
        $i = $i + 1    
$outObject."Number" = $i
        $outObject."Received" = $email.Received.ToLocalTime()
        $outObject."Sender Address" = $email.SenderAddress
        $outObject."Subject" = $email."Subject"
    
        $collection += $outObject
}
$collection | Out-Host
Do {
        Try {
            $num = $true
            [int]$selectedinput = Read-host "Select an email by typing its number"
        }
        Catch {$num = $false}
    }
    Until (($selectedinput -gt 0 -and $selectedinput -le $collection.count) -and $num -eq $true)
$selectedemail = $collection[[int]$selectedinput - 1]
$selectedemail
$sendere = $selectedemail.'Sender Address'
    $subject = $selectedemail.Subject
    
}
$today = get-date -Format yyyy/MM/dd_HH-mm-ss
$searchname = "Exchange_"+$today
$tempquery = "Subject:""" + $subject + """" + " AND "+ "received>=""" + $startdate + """ AND " + "From=" + $sendere
$query = $tempquery
Write-Host "Your search query:"`n$tempquery`n
New-ComplianceSearch -Name $searchname -ExchangeLocation all -ContentMatchQuery $query | Format-List @{Name = "Compliance Search Name"; Expression = {$_.Name}}
Start-ComplianceSearch -Identity $searchname
Write-Host "Search request has been created"
$k = 0
Do {
    $k = $k + 1
    Write-Host "." -NoNewline -ForegroundColor Cyan
    Start-Sleep -Seconds 1
}
While ((Get-ComplianceSearch -Identity $searchname).Status -ne "Completed")
Write-Host `n"Search complete within "$k" seconds." -ForegroundColor Green
$compliancesearch = Get-ComplianceSearch -Identity $searchname
$foundresults = $compliancesearch.SuccessResults
$array = $foundresults.Split([Environment]::NewLine,[System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object {
    $_ -notlike "*Item count: 0*"
    }
Write-host `n"Search results:"
If ($array.Count -eq 0) {
    Write-Host "0 items have been found, try another queue."
}
Else {
    $array
$answer = Read-Host `n"Would you like to export those messages from the list above and then purge them ( y / n )?"
while("y","n" -notcontains $answer) {$answer = Read-Host "Would you like to export those messages from the list above and then purge them ( y / n )?"}
If ($answer -eq 'y') {
New-ComplianceSearchAction -SearchName $searchname -Export -ExchangeArchiveFormat SinglePst -Format FXStream | Format-List Name,SearchName,Action,RunBy
Write-Host "Export request has been created"
$k = 0
        Do {
            $k = $k + 1
            Write-Host "." -NoNewline -ForegroundColor Cyan
            Start-Sleep -Seconds 1
        }
        While ((Get-ComplianceSearchAction -Identity $searchname"_Export").Status -ne "Completed")
Write-Host `n"Export complete within "$k" seconds." -ForegroundColor Green
        
New-ComplianceSearchAction -SearchName $searchname -Purge -PurgeType HardDelete | Format-List Name,SearchName,Action,RunBy
    
        Write-Host "Purge request has been created"
$k = 0
        Do {
            $k = $k + 1
            Write-Host "." -NoNewline -ForegroundColor Cyan
            Start-Sleep -Seconds 1
        }
        While ((Get-ComplianceSearchAction -Identity $searchname"_Purge").Status -ne "Completed")
Write-Host `n"Purge complete within "$k" seconds." -ForegroundColor Green
Write-Host `n"To download the exported data, please go to https://protection.office.com/contentsearchbeta?ContentOnly=1 / Export tab."
       
    }
}
Write-Host `n"Clearing only Exchange_date_time Search, Export, and Purge compliance search results older than" $retentiondays "days."
Get-ComplianceSearch | Where-Object {
    $_.JobEndTime -lt (Get-Date).AddDays(-$retentiondays) -and
    $_.Status -eq "Completed" -and
    $_.Name -match '\d\d\d\d/\d\d/\d\d_\d\d-\d\d-\d\d'
} | Remove-ComplianceSearch -Confirm:$false
Get-ComplianceSearchAction | Where-Object {
    $_.JobEndTime -lt (Get-Date).AddDays(-$retentiondays) -and
    $_.Status -eq "Completed" -and
    (
        $_.Name -match 'Exchange_\d\d\d\d/\d\d/\d\d_\d\d-\d\d-\d\d_Purge' -or
        $_.Name -match 'Exchange_\d\d\d\d/\d\d/\d\d_\d\d-\d\d-\d\d_Export'
    )
} | Remove-ComplianceSearchAction -Confirm:$false
Write-Host `n"Complete"`n -ForegroundColor Green
Write-Host "Total Time: $($elapsed.Elapsed.ToString())"
Remove-Variable * -ErrorAction SilentlyContinue
