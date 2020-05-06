Add-PSSnapin Microsoft.SharePoint.PowerShell

#Removing Existing Search Service Applications

$SSAS = Get-SPEnterpriseSearchServiceApplication

ForEach ($SSA in $SSAS)

{

$SSA | FT Name, ID, ApplicationPool

$Response = Read-Host -Prompt “Would you like to remove the above Search Service Application and all associated search data? Press Y or N”

IF ($Response -eq “y”)

{

Write-Host ‘Removing SSA’ $SSA.Id

$AllProxies = Get-SPEnterpriseSearchServiceApplicationProxy

$Proxy = $AllProxies | ?{$_.GetSearchServiceApplicationInfo().SearchServiceApplicationID -eq $SSA.Id}

Remove-SPEnterpriseSearchServiceApplicationProxy -Identity $Proxy

Remove-SPEnterpriseSearchServiceApplication -Identity $SSA -RemoveData

}

ELSE

{

Write-Host “Skipping SSA Removal”

}

}

#Stop Search Service Instances

$SSIS = Get-SPEnterpriseSearchServiceInstance

ForEach ($SSI in $SSIS)

{

$SSI

$Response = Read-Host -Prompt “Would you like to stop the above service? Press Y or N”

IF ($Response -eq “Y”)

{

Write-Host ‘Stopping’ $SSIS.Service ‘on’ $SSIS.Server

Stop-SPEnterpriseSearchServiceInstance -Identity $SSI

}

}

#Stop Search Query and Site Settings Service Instances

$SQSSSIS = Get-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance

ForEach ($SQSSSI in $SQSSSIS)

{

$SQSSSI

$Response = Read-Host -Prompt “Would you like to stop the above service? Press Y or N”

IF ($Response -eq “Y”)

{

Write-Host ‘Stopping’ $SQSSsI.Service ‘on’ $SQSsSI.Server

Stop-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance -Identity $SQSsSI

}

}

#Check Service Status

Get-SPEnterpriseSearchServiceInstance | FT Server, Status

Get-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance | FT Server, Status
