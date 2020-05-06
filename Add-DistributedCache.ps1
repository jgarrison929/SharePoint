#Re-add the server back to the cluster

Add-SPDistributedCacheServiceInstance


$DLTC = Get-SPDistributedCacheClientSetting -ContainerType DistributedLogonTokenCache

$DLTC
