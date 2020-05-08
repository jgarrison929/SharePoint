### SharePoint New Project Site Script
### Created by: Joshua Garrison
### Change Log: {2017-10-02} Added Function to remove Access Request on all sites after deployment 
/* 
* Dear maintainer:
* 
* Once you are done trying to 'optimize' this routine,
* and have realized what a terrible mistake that was,
* please increment the following counter as a warning
* to the next guy:
* 
* total_hours_wasted_here = 42
 */

if ($ver.Version.Major -gt 1) {$host.Runspace.ThreadOptions = "ReuseThread"} 
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) 
{
     Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

. .\New-SPGroup.ps1
. .\Add-Sppermission.ps1

### Instance variables: Set ParentURL,web, jobNumber, jobName, and PM ###
### Save New Project.ps1 and execute as Administrator from SharePoint Management Shell ###



$jobNumber = Read-Host "Enter Job Number (E.G. 52.8408)"
$jobName = Read-Host "Enter Job Name (E.G. Visalia RFB)"
$PM = Read-Host "Enter the Project Manager AD Account UPN"
$SPURL = Read-Host "Enter the SharePoint Web Application URL"


### Instance Variable End ###

#########################################################################################################################################################

### DO NOT EDIT THESE ###

$parentURL = $SPURL
$web = Get-SPWeb $parentURL
$subSiteUrl = $jobNumber -replace '[.]'
$siteName = $jobNumber + " " + $jobName
$LSCIT = $web.groups["LSC-IT"]
$LSCRM = $web.groups["LSC-RM"]
$NESMUPDATE = $web.groups["NESM-Update"]
$NESMREAD = $web.groups["NESM-READ"]



### DO NOT EDIT THESE ###



### Create subsite ###
# If we use the -template property when we use New-SPWeb, the script fails. That's why we need to apply template
# after the subsite is created
Write-Host "Creating Website for" $siteName -ForegroundColor Green
$mainurl = $parentURL + $subSiteUrl
$template = $web.GetAvailableWebTemplates(1033) | Where-Object {$_.Title -eq "MasterTemplate20180906"}
$web = New-SPWeb $mainurl -Language 1033 -Name $siteName -UniquePermissions
Write-Host "Applying Template" -ForegroundColor Green
$web.ApplyWebTemplate($template.Name)

###Break Inheritance on All Libraries and Lists###
Write-Host "Breaking  Inheritance on All Document Libraries and Lists" -ForegroundColor Green
foreach($l in $web.lists)
 {
 $web.lists.BreakRoleInheritance($true,$true)
 }
    
# Create JobNumber + NESM-Admins Group ###
Write-Host "Creating ${jobNumber} NESM-Admin Group" -ForegroundColor Yellow
$groupName1 = $jobNumber + " " + "NESM-Admin"
$ownerName = "$PM"
$memberName = "$PM"
$description = "${jobNumber} NESM-Admin Group"
New-SPGroup -web $mainurl -GroupName "$groupName1" -OwnerName $PM -MemberName $PM -description $description
	
	
### Create JobNumber + NESM-Update Group ###
Write-Host "Creating ${jobNumber} NESM-Update Group" -ForegroundColor Yellow
$groupName2 = $jobNumber + " " + "NESM-Update"
$ownerName = "$PM"
$memberName = "$PM"
$description = "${jobNumber} NESM-Update Group"   
New-SPGroup -web $mainurl -GroupName "$groupName2" -OwnerName $PM -MemberName $pm -description $description

	
### Create JobNumber + NESM-Read Group ###
Write-Host "Creating ${jobNumber} NESM-Read Group" -ForegroundColor Yellow
$groupName3 = $jobNumber + " " + "NESM-Read"
$ownerName = "$PM"
$memberName = "$PM"
$description = "${jobNumber} NESM-Read Group"
New-SPGroup -web $mainurl -GroupName "$groupName3" -OwnerName $PM -MemberName $pm -description $description

		
### Create JobNumber + Subs-Update Group ###
Write-Host "Creating ${jobNumber} Subs-Update Group" -ForegroundColor Yellow
$groupName4 = $jobNumber + " " + "Subs-Update"
$ownerName = "$PM"
$memberName = "$PM"
$description = "${jobNumber} Subs-Update Group"
New-SPGroup -web $mainurl -GroupName "$groupName4" -OwnerName $PM -MemberName $pm -description $description

	
### Create JobNumber + Subs-Read Group ###
Write-Host "Creating ${jobNumber} Subs-Read Group" -ForegroundColor Yellow
$groupName5 = $jobNumber + " " + "Subs-Read"
$ownerName = "$PM"
$memberName = "$PM"
$description = "${jobNumber} Subs-Read Group"
New-SPGroup -web $mainurl -GroupName "$groupName5" -OwnerName $PM -MemberName $pm -description $description	

	
### Create JobNumber + Owners-Update Group ###
Write-Host "Creating ${jobNumber} Owners-Update Group" -ForegroundColor Yellow
$groupName6 = $jobNumber + " " + "Owners-Update"
$ownerName = "$PM"
$memberName = "$PM"
$description = "${jobNumber} Owners-Update Group"
New-SPGroup -web $mainurl -GroupName "$groupName6" -OwnerName $PM -MemberName $pm -description $description
	

### Create JobNumber + Owners-Read Group ###
Write-Host "Creating ${jobNumber} Owners-Read Group" -ForegroundColor Yellow
$groupName7 = $jobNumber + " " + "Owners-Read"
$ownerName = "$PM"
$memberName = "$PM"
$description = "${jobNumber} Owners-Read Group"
New-SPGroup -web $mainurl -GroupName "$groupName7" -OwnerName $PM -MemberName $pm -description $description

### Set permission on group JobNumber + NESM-Admin ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$group = $web.Groups["$groupName1"]
$role = $web.RoleDefinitions["Full control"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($group)
$RoleAssignment.RoleDefinitionBindings.Add($role)


# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()


### Set permission on group JobNumber + NESM-Update ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$group = $web.Groups["$groupName2"]
$role = $web.RoleDefinitions["Contribute"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($group)
$RoleAssignment.RoleDefinitionBindings.Add($role)


# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()	
	

### Set permission on group JobNumber + NESM-Read ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$group = $web.Groups["$groupName3"]
$role = $web.RoleDefinitions["Read"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($group)
$RoleAssignment.RoleDefinitionBindings.Add($role)
# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()	
	
	
### Set permission on group JobNumber + Subs-Update ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$group = $web.Groups["$groupName4"]
$role = $web.RoleDefinitions["Read"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($group)
$RoleAssignment.RoleDefinitionBindings.Add($role)
# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()
	
### Set permission on group JobNumber + Subs-Read ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$group = $web.Groups["$groupName5"]
$role = $web.RoleDefinitions["Read"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($group)
$RoleAssignment.RoleDefinitionBindings.Add($role)
# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()	


### Set permission on group JobNumber + Owner-Update ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$group = $web.Groups["$groupName6"]
$role = $web.RoleDefinitions["Read"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($group)
$RoleAssignment.RoleDefinitionBindings.Add($role)
# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()


### Set permission on group JobNumber + Owner-Read ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$group = $web.Groups["$groupName7"]
$role = $web.RoleDefinitions["Read"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($group)
$RoleAssignment.RoleDefinitionBindings.Add($role)
# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()



### Set permission on group JobNumber + NESM-Update ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$role = $web.RoleDefinitions["Contribute"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($NESMUPDATE)
$RoleAssignment.RoleDefinitionBindings.Add($role)

# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()


### Set permission on group JobNumber + LSC-IT ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$role = $web.RoleDefinitions["Full Control"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($LSCIT)
$RoleAssignment.RoleDefinitionBindings.Add($role)

# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()	
	

### Set permission on group JobNumber + NESM-Read ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$role = $web.RoleDefinitions["Read"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($NESMREAD)
$RoleAssignment.RoleDefinitionBindings.Add($role)

# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()	
	
	
### Set permission on group JobNumber + LSC-RM ###
Write-Host "  Setting permissions on group" -ForegroundColor Yellow
$role = $web.RoleDefinitions["Read"]
$RoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($LSCRM)
$RoleAssignment.RoleDefinitionBindings.Add($role)

# the following could be for a web, list or item object
$web.RoleAssignments.Add($RoleAssignment)
$web.Update()
$web.Dispose()
	
<##	
	
###Get All Libraries and Lists in $web and Grant Permissions ###
Write-Host " Setting permissions on all Lists and Libraries for $groupName1"
foreach ($l in $web.lists)
 {Add-SPPermissionToListGroup $mainurl "$l" "$groupName1"  "Full Control"}
	
	
###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on all Lists and Libraries for $groupName2"
foreach ($l in $web.lists)
 {Add-SPPermissionToListGroup $mainurl "$l" "$groupName2"  "Contribute"}

 
###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on all Lists and Libraries for $groupName3"
foreach ($l in $web.lists)
 {Add-SPPermissionToListGroup $mainurl "$l" "$groupName3"  "Read"}

 
###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on RFI's for $groupName4"
Add-SPPermissionToListGroup $mainurl "RFI's" "$groupName4"  "Contribute"


###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on RFI Log for $groupName4"
Add-SPPermissionToListGroup $mainurl "RFI Log" "$groupName4"  "Contribute"
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Submittals for $groupName4"
Add-SPPermissionToListGroup $mainurl "Submittals" "$groupName4"  "Contribute"
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Submittal Log for $groupName4"
Add-SPPermissionToListGroup $mainurl "Submittal Log" "$groupName4"  "Contribute" 
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Daily Reports for $groupName4"
Add-SPPermissionToListGroup $mainurl "Daily Reports" "$groupName4"  "Contribute"


###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on RFI's for $groupName5"
Add-SPPermissionToListGroup $mainurl "RFI's" "$groupName5"  "Read"


###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on RFI Log for $groupName5"
Add-SPPermissionToListGroup $mainurl "RFI Log" "$groupName5"  "Read"
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Submittals for $groupName5"
Add-SPPermissionToListGroup $mainurl "Submittals" "$groupName5"  "Read"
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Submittal Log for $groupName5"
Add-SPPermissionToListGroup $mainurl "Submittal Log" "$groupName5"  "Read" 
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Daily Reports for $groupName5"
Add-SPPermissionToListGroup $mainurl "Daily Reports" "$groupName5"  "Read"
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on RFI's for $groupName6"
Add-SPPermissionToListGroup $mainurl "RFI's" "$groupName6"  "Contribute"
	

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on RFI Log for $groupName6"
Add-SPPermissionToListGroup $mainurl "RFI Log" "$groupName6"  "Contribute"
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Submittals for $groupName6"
Add-SPPermissionToListGroup $mainurl "Submittals" "$groupName6"  "Contribute"
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Submittal Log for $groupName6"
Add-SPPermissionToListGroup $mainurl "Submittal Log" "$groupName6"  "Contribute" 
 

###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on Daily Reports for $groupName6"
Add-SPPermissionToListGroup $mainurl "Daily Reports" "$groupName6"  "Contribute"	 
 

###Grant Read Permissions to RFI's for $groupName7 ###
Write-Host " Setting permissions on RFI's for $groupName7"
Add-SPPermissionToListGroup $mainurl "RFI's" "$groupName7"  "Read"
	

###Grant Read Permissions to RFI Log for $groupName7 ###
Write-Host " Setting permissions on RFI Log for $groupName7"
Add-SPPermissionToListGroup $mainurl "RFI Log" "$groupName7"  "Read"
 

###Grant Read Permissions to Submittals for $groupName7 ###
Write-Host " Setting permissions on Submittals for $groupName7"
Add-SPPermissionToListGroup $mainurl "Submittals" "$groupName7"  "Read"
 

###Grant Read Permissions to Submittal Log for $groupName7 ###
Write-Host " Setting permissions on Submittal Log for $groupName7"
Add-SPPermissionToListGroup $mainurl "Submittal Log" "$groupName7"  "Read" 
 

###Grant Read Permissions to Daily Reports for $groupName7 ###
Write-Host " Setting permissions on Daily Reports for $groupName7"
Add-SPPermissionToListGroup $mainurl "Daily Reports" "$groupName7"  "Read"	 

##>

###Get All Libraries and Lists in $web and Grant Permissions ###
Write-Host " Setting permissions on all Lists and Libraries for $LSCIT"
foreach ($l in $web.lists)
 {Add-SPPermissionToListGroup $mainurl "$l" "$LSCIT"  "Full Control"}
	
	
###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on all Lists and Libraries for $NESMUPDATE"
foreach ($l in $web.lists)
 {Add-SPPermissionToListGroup $mainurl "$l" "$NESMUPDATE"  "Contribute"}

 
###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on all Lists and Libraries for $NESMREAD"
foreach ($l in $web.lists)
 {Add-SPPermissionToListGroup $mainurl "$l" "$NESMREAD"  "Read"}
 
 
 ###Get All Libraries and Lists in $web ###
Write-Host " Setting permissions on all Lists and Libraries for $LSCRM"
foreach ($l in $web.lists)
 {Add-SPPermissionToListGroup $mainurl "$l" "$LSCRM"  "Read"}

 
 
#Get all sites
$WebsColl = Get-SPWebApplication $parentURL | Get-SPSite -Limit All | Get-SPWeb -Limit All
 
ForEach ($web in $WebsColl)
    {
        if($web.RequestAccessEnabled -and $web.Permissions.Inherited -eq $false)
        {
            #Disable access request
            $web.RequestAccessEmail=""
            $web.Update()
            write-host "Access request disabled at site:"$web.URL
        }
    }

<##

#Grant Access to Project Site for All NESM-Update Users
$siteCollUrl = $mainurl

$users = Get-Spweb $parenturl | Select -ExpandProperty SiteGroups | Where {$_.Name -EQ "NESM-Update"} | Select -ExpandProperty Users | Select  userlogin 
$user1 = $users | foreach-object{$_.userlogin.trimstart("i:0#.w|")}

$web = Get-SPWeb -identity $siteCollUrl
$group = $web.SiteGroups["$groupname2"]
foreach ($User in $user1) { 
    $web.EnsureUser($User)	
    New-SPUser $User -web $siteCollURL -Group $group	
} 

##>

