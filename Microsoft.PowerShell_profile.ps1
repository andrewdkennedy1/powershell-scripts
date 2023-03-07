Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -LiteralPath $PSCommandPath
$startingLoc = Get-Location
Set-Location $scriptDir
$startingDir = [System.Environment]::CurrentDirectory
[System.Environment]::CurrentDirectory = $scriptDir

function get-info {
    [CmdletBinding()]
    param([string]$MyParam)
    try {
        $user = Get-ADUser -Filter "name -like '*$MyParam*'" -Properties UserPrincipalName,mail | Select-Object UserPrincipalName,mail
        $computer = Get-ADComputer -Filter "description -like '*$MyParam*'" -Properties DNSHOSTNAME,WhenChanged,description | Select-Object DNSHOSTNAME,WhenChanged,description
        $output = @($user, $computer)
        Write-Output $output | Sort-Object -Property WhenChanged | Format-Table -Property UserPrincipalName,mail,DNSHOSTNAME,WhenChanged,description
        $pingstatus = $computer | Select-Object -ExpandProperty DNSHOSTNAME
        foreach ($c in $pingstatus) {
            $ping = Test-Connection -ComputerName $c -Count 1 -Quiet
            if ($ping) {
                Write-Output "$c is online"
            } else {
                Write-Output "$c is offline"
            }
        }
    } catch {
        Write-Error $_.Exception
    }
}

function Get-Printers {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$MyParam
    )
    try {
        $gp1 = Get-WMIObject -Class Win32_Printer -Computer bpc-vm-printers.ad.bortonfruit.com | Select Name,DriverName,PortName,Shared,ShareName
        Write-Output $gp1 | Where-Object {$_.name -Match $MyParam} | ft -auto
    } catch {
        Write-Error $_.Exception
    }
}

function uptime {
    try {
        Get-WmiObject win32_operatingsystem | select csname, @{LABEL='LastBootUpTime'; EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}
    } catch {
        Write-Error $_.Exception
    }
}

function Start-ElevatedScript {
    Param($Script)
    $script = Join-Path $PWD $Script
    Start-Process powershell -ArgumentList "-NoExit","-File","$script" -Verb runAs
}

Set-Alias elevate Start-ElevatedScript -Scope Global


Set-Alias who get-info
Set-Alias pp Get-Printers
Set-Location "c:\scripts"
