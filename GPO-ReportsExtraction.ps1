<#
Script - This Script will help you Export all the Group Policy in HTML format.
Created - Feb 4, 2021
Author - Binu Balan
#>

cls
$host.ui.RawUI.WindowTitle = "APPU - DHCP Report Extract"
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host " "
Write-host "                 _    ____  ____  _   _ "
Write-host "                / \  |  _ \|  _ \| | | |"
Write-host "               / _ \ | |_) | |_) | | | |"
Write-host "              / ___ \|  __/|  __/| |_| |"
Write-host "             /_/   \_\_|   |_|    \___/ "
Write-Host " " 
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host " " 
Write-host "	           .__." -ForegroundColor Green
Write-host "                   (oo)____" -ForegroundColor Green
Write-host "                   (__)    )\" -ForegroundColor Green
Write-host "                      ll--ll '" -ForegroundColor Green
Write-Host "               SCRIPT BY BINU BALAN               " -ForegroundColor DarkRed -BackgroundColor White 
Write-Host "<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ForegroundColor Cyan
Write-Host ""  
Write-Host ""

Start-Sleep -Seconds 3

Import-Module GroupPolicy
$ReportPath = "C:\GPOExtractions"
$Checkreportpath = Test-Path -Path "C:\GPOExtractions"
if($Checkreportpath -eq $false){
    New-Item -ItemType Directory -Path C:\ -Name GPOExtractions | Out-Null
}

Function ExportGPO {
$AllGroupPolicy = Get-GPO -All | select DisplayName,ID
    ForEach ($eachgpo in $AllGroupPolicy){
        $gponame = $eachgpo.DisplayName
        $gpoguid = $eachgpo.ID
        Write-Host "Exporting GPO $gponame " -NoNewline
        Get-GPOReport -Guid $gpoguid -ReportType Html -Path $ReportPath\$gponame.html
        Write-Host "[ " -NoNewline -ForegroundColor Yellow
        Write-host "Completed" -NoNewline -ForegroundColor Green
        Write-host " ]" -ForegroundColor Yellow

    }

}

ExportGPO

Write-Host "=================== Script Ended ==================="