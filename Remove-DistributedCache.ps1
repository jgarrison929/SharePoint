#Stopping the service on local host

Stop-SPDistributedCacheServiceInstance -Graceful

#Removing the service from SharePoint on local host.

Remove-SPDistributedCacheServiceInstance

#Cleanup left over pieces from SharePoint

$instanceName =”SPDistributedCacheService Name=AppFabricCachingService”

$serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq $env:computername}

$serviceInstance.delete()
