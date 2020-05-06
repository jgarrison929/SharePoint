Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue
 
#Get all SharePoint Servers
$Servers = Get-SPServer | ? {$_.Role -ne "Invalid"} | Select -ExpandProperty Address
 
#Iterate through each server and reset SharePoint config cache
Invoke-Command -ComputerName $Servers -ScriptBlock {
try {
        Write-Host "$env:COMPUTERNAME - Stopping timer service"
        Stop-Service SPTimerV4
 
        #Get Config Cache Folder
        $ConfigDbId = [Guid](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\16.0\Secure\ConfigDB' -Name Id).Id #Path to the '15 hive' ConfigDB in the registry
        $CacheFolder = Join-Path -Path ([Environment]::GetFolderPath("CommonApplicationData")) -ChildPath "Microsoft\SharePoint\Config\$ConfigDbId"
 
        Write-Host "$env:COMPUTERNAME - Clearing cache folder $CacheFolder"
        #Delete all XML Files
        Get-ChildItem "$CacheFolder\*" -Filter *.xml | Remove-Item
 
        Write-Host "$env:COMPUTERNAME - Resetting cache ini file"
        $CacheIni = Get-Item "$CacheFolder\Cache.ini"
        Set-Content -Path $CacheIni -Value "1"
    }
finally{
        Write-Host "$env:COMPUTERNAME - Starting timer service"
        Start-Service SPTimerV4
    }
}
