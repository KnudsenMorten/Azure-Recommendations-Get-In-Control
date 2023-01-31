#Requires -Version 5.0
<#
    .SYNOPSIS
    
    In this script, I will demonstrate, how you can extract security recommendations from Microsoft Defender for Cloud
    using Azure Resource Graph - delivering a horizontal cross-subscriptions, workload overview. 
    Data will automatically be exported into a Excel spreadsheet delivering 19 Excel tables and 16 pivot tables.

    Information can be used to detect deviations from best practice / desired state - covering
    * Getting-in-control with workloads in tenant/management group (storage, network, app services, containers, etc.) 
      where we are not in control according to security best practice / desired state
    * Getting-in-control with subscriptions, where environment are not configured according to security best practice / desired state
    * Get-in-control with role assignments in tenant / management group / subscription / resource group.
    * Get detailed information about role assignments on user / service-principal-level, based on direct assignment and inheritance
    * Get detailed insight about users / service-principal-level, based on group membership - both direct and inheritance.

    .NOTES
    VERSION: 2301

    .COPYRIGHT
    @mortenknudsendk on Twitter
    Blog: https://mortenknudsen.net
    
    .LICENSE
    Licensed under the MIT license.

    .WARRANTY
    Use at your own risk, no warranty given!
#>


#------------------------------------------------------------------------------------------------------------
# Functions
#------------------------------------------------------------------------------------------------------------
Function AZ_Find_Subscriptions_in_Tenant_With_Subscription_Exclusions
{
    Write-Output ""
    Write-Output "Finding all subscriptions in scope .... please Wait !"

    $global:Query_Exclude = @()
    $Subscriptions_in_Scope = @()
    $pageSize = 1000
    $iteration = 0

    ForEach ($Sub in $global:Exclude_Subscriptions)
        {
            $global:Query_Exclude    += "| where (subscriptionId !~ `"$Sub`")"
        }

    $searchParams = @{
                        Query = "ResourceContainers `
                                | where type =~ 'microsoft.resources/subscriptions' `
                                | extend status = properties.state `
                                $global:Query_Exclude
                                | project id, subscriptionId, name, status | order by id, subscriptionId desc" 
                        First = $pageSize
                        }

    $results = do {
        $iteration += 1
        $pageResults = Search-AzGraph -ManagementGroup $Global:ManagementGroupScope @searchParams
        $searchParams.Skip += $pageResults.Count
        $Subscriptions_in_Scope += $pageResults
    } while ($pageResults.Count -eq $pageSize)

    $Global:Subscriptions_in_Scope = $Subscriptions_in_Scope

    # Output
    $Global:Subscriptions_in_Scope
}


#------------------------------------------------------------------------------------------------------------
# Variables
#------------------------------------------------------------------------------------------------------------

# Scope (MG) | You can define the scope for the targetting, supporting management groups or tenant root id (all subs)
$Global:ManagementGroupScope                                = "f0fa27a0-8e7c-4f63-9a77-ec94786b7c9e" # can mg e.g. mg-company or AAD Id (=Tenant Root Id)

# Exclude list | You can exclude certain subs, resource groups, resources, if you don't want to have them as part of the scope
$global:Exclude_Subscriptions                               = @("xxxxxxxxxxxxxxxxxxxxxx") # for example platform-connectivity
$global:Exclude_ResourceGroups                              = @()
$global:Exclude_Resource                                    = @()
$global:Exclude_Resource_Contains                           = @()
$global:Exclude_Resource_Startswith                         = @()
$global:Exclude_Resource_Endswith                           = @()

# Content file
$HelpContentFile                                            = "C:\SCRIPTS\Azure-Recommendations-Get-In-Control\Content.csv"

# OutputFile
$FileOutput                                                 = "C:\SCRIPTS\Azure-Recommendations-Get-In-Control\Azure_Recommendations_Get-in-Control.xlsx"


#------------------------------------------------------------------------------------------------------------
# Connect to Azure & Azure AD
#------------------------------------------------------------------------------------------------------------
Connect-AzAccount
Connect-AzureAD


#--------------------------------------------------------------------------------------------------------
# Powershell Modules
#--------------------------------------------------------------------------------------------------------
<#
Install-module Az
Install-module Az.ResourceGraph
Install-module ImportExcel
#>

#------------------------------------------------------------------------------------------------------------
# File Check for Content.csv
#------------------------------------------------------------------------------------------------------------

If (!(Test-Path $HelpContentFile))
    {
        Write-Output "Content-file $($HelpContentFile) was NOT found !!"
        Break
    } 


#--------------------------------------------------------------------------------------------------------
# Scope - subscriptions - where to look for Azure Role Assignments
#--------------------------------------------------------------------------------------------------------

# Retrieving data using Azure ARG
AZ_Find_Subscriptions_in_Tenant_With_Subscription_Exclusions

# Filter - only lists ENABLED subscriptions
$Global:Subscriptions_in_Scope = $Global:Subscriptions_in_Scope | Where-Object {$_.status -eq "Enabled"}

Write-Output "Scope (target)"

$Global:Subscriptions_in_Scope



###############################
# (A1) Azure RBAC Role Assignments
###############################

    $RBAC_RoleAssignments = @()
    ForEach ($Subscription in $Global:Subscriptions_in_Scope)
        {
            # Get Azure RBAC Role Assignments
            Write-output ""
            $AzContext = Set-AzContext -SubscriptionId $Subscription.subscriptionId -Force

            $SubscriptionName  = $Subscription.Name
            $SubscriptionId    = $Subscription.subscriptionId  
 
            # Getting information about Role Assignments for choosen subscription
            Write-output "[$($Subscription.Name)] - Getting Role Assignments"
        
            $RBAC_Roles = Get-AzRoleAssignment -WarningAction SilentlyContinue
  
             foreach ($Role in $RBAC_Roles)
                {
                    $AccObj = New-Object -TypeName PSObject

                    If ($Role.ObjectType -eq "Unknown")
                        {
                            Write-Output "    [$($Role.RoleDefinitionName)] - Unknown (clean-up needed)"
                        }
                    Else
                        {
                            Write-Output "    [$($Role.RoleDefinitionName)] - $($Role.DisplayName)"
                        }

                    # Role Assignment info
                    $RoleAssignmentName = $role.RoleAssignmentName
                    $RoleAssignmentId   = $role.RoleAssignmentId

                    # Role Assignement Scope - Direct on Sub or Inheritance

                    # Inheritance
                    If ($Role.Scope -eq "/subscriptions/$($SubscriptionId)")
                        {
                            $Scope_Delegation = "Direct_SUB"
                            $Scope            = $Role.Scope
                        }
                    ElseIf ($Role.Scope -like "/subscriptions/$($SubscriptionId)/*")
                        {
                            $Scope_Delegation = "Direct_RG"
                            $Scope            = $Role.Scope
                        }
                    ElseIf ($Role.Scope -like "/providers/Microsoft.Management/managementGroups*")
                        {
                            $Scope_Delegation = "Inheritance_MG"
                            $Scope            = ($Role.Scope.split("/")[4])
                        }


                    # Direct Account Role Assignment (not group)
                    If ( ($Role.ObjectType -ne "Group") -and ($Role.ObjectType -ne "Unknown") )
                        {
                            # Delegated Account (direct)
                            $AccInfo           = Get-AzureADObjectByObjectId -ObjectIds $Role.ObjectId
                            $DisplayName       = $AccInfo.DisplayName
                            $UserPrincipalName = $AccInfo.UserPrincipalName
                            $ObjectId          = $AccInfo.ObjectId
                            $ObjectType        = $AccInfo.ObjectType
                            $AccEnabled        = $AccInfo.AccountEnabled

                            $AccObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DisplayName
                            $AccObj | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $UserPrincipalName
                            $AccObj | Add-Member -MemberType NoteProperty -Name ObjectType -value $ObjectType
                            $AccObj | Add-Member -MemberType NoteProperty -Name ObjectId -value $ObjectId
                            $AccObj | Add-Member -MemberType NoteProperty -Name AccEnabled -value $AccEnabled
                            $AccObj | Add-Member -MemberType NoteProperty -Name RBAC_Delegation_Type -value "Direct"
                            $AccObj | Add-Member -MemberType NoteProperty -Name RBAC_GroupName -value "N/A"
                        }

                    ElseIf ( ($Role.ObjectType -eq "Unknown") )
                        {
                            # Delegated Account (direct)
                            $DisplayName       = ""
                            $UserPrincipalName = ""
                            $ObjectId          = $Role.ObjectId
                            $ObjectType        = "Unknown"
                            $AccEnabled        = "NotFound"

                            $AccObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DisplayName
                            $AccObj | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $UserPrincipalName
                            $AccObj | Add-Member -MemberType NoteProperty -Name ObjectType -value $ObjectType
                            $AccObj | Add-Member -MemberType NoteProperty -Name ObjectId -value $ObjectId
                            $AccObj | Add-Member -MemberType NoteProperty -Name AccEnabled -value $AccEnabled
                            $AccObj | Add-Member -MemberType NoteProperty -Name RBAC_Delegation_Type -value "Direct"
                            $AccObj | Add-Member -MemberType NoteProperty -Name RBAC_GroupName -value "N/A"
                        }

                    # Delegated Account (indirect due to group)

                    ElseIf ($Role.ObjectType -eq "Group")
                        {
                            $AccObj = @()
                            $IndirectMembers = Get-AzureADGroupMember -ObjectId $Role.ObjectId -All $true | Select-Object DisplayName, SignInName, ObjectId, ObjectType
                            ForEach ($Member in $IndirectMembers)
                                {
                                    $AccInfo = Get-AzureADObjectByObjectId -ObjectIds $Role.ObjectId
                                    $DisplayName       = $AccInfo.DisplayName
                                    $UserPrincipalName = $AccInfo.UserPrincipalName
                                    $ObjectId          = $AccInfo.ObjectId
                                    $ObjectType        = $AccInfo.ObjectType
                                    $AccEnabled        = $AccInfo.AccountEnabled

                                    $Indirectobj = New-Object -TypeName PSObject
                                    $Indirectobj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DisplayName
                                    $Indirectobj | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $UserPrincipalName
                                    $Indirectobj | Add-Member -MemberType NoteProperty -Name ObjectType -value $ObjectType
                                    $Indirectobj | Add-Member -MemberType NoteProperty -Name ObjectId -value $ObjectId
                                    $Indirectobj | Add-Member -MemberType NoteProperty -Name AccEnabled -value $AccEnabled
                                    $Indirectobj | Add-Member -MemberType NoteProperty -Name RBAC_Delegation_Type -value "Group_inheritance"
                                    $Indirectobj | Add-Member -MemberType NoteProperty -Name RBAC_GroupName -value $Role.DisplayName

                                    $AccObj += $Indirectobj
                                }
                        }

                    # Role Delegation
                    $RoleDefName        = $role.RoleDefinitionName
                    $RoleDefId          = $role.RoleDefinitionId

                    # Subscription Info
                    $SubscriptionName   = $AzContext.Subscription.Name
                    $SubscriptionId     = $AzContext.Subscription.SubscriptionId
 
                    # Checking for Custom Role
                    $CustomRolesIfAny   = Get-AzRoleDefinition -Name $RoleDefName
                    $CustomRole         = $CustomRolesIfAny.IsCustom

                    # Make an array with the result - used for reporting in Excel
                    ForEach ($Entry in $AccObj)
                        {
                            $Roleobj = New-Object -TypeName PSObject
                            $Roleobj | Add-Member -MemberType NoteProperty -Name SubscriptionName -value $SubscriptionName
                            $Roleobj | Add-Member -MemberType NoteProperty -Name SubscriptionId -value $SubscriptionId         

                            $Roleobj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $Entry.DisplayName
                            $Roleobj | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $Entry.UserPrincipalName
                            $Roleobj | Add-Member -MemberType NoteProperty -Name ObjectType -value $Entry.ObjectType
                            $Roleobj | Add-Member -MemberType NoteProperty -Name ObjectId -value $Entry.ObjectId
                            $Roleobj | Add-Member -MemberType NoteProperty -Name AccEnabled -value $Entry.AccEnabled
                            $Roleobj | Add-Member -MemberType NoteProperty -Name RBAC_Delegation_Type -value $Entry.RBAC_Delegation_Type
                            $Roleobj | Add-Member -MemberType NoteProperty -Name RBAC_GroupName -value $Entry.RBAC_GroupName

                            $Roleobj | Add-Member -MemberType NoteProperty -Name RoleDefinitionName -value $RoleDefName
                            $Roleobj | Add-Member -MemberType NoteProperty -Name CustomRole -value $CustomRole
                            $Roleobj | Add-Member -MemberType NoteProperty -Name Scope -value $Scope
                            $Roleobj | Add-Member -MemberType NoteProperty -Name Scope_Delegation -value $Scope_Delegation

                            $RBAC_RoleAssignments += $Roleobj
                        }
        }
    }



<#  EXCLUDED SINCE IT IS VERY SLOW IN A LARGE ENTERPRISE SETUP !!!
#########################################################
# (A2) MDC | Recommendations with SubAssessments (VERY SLOW)
#########################################################


    $MDC_Recommendations = @()
    $pageSize = 1000
    $iteration = 0
    $searchParams = @{
                        Query = "SecurityResources `
                                | where type == 'microsoft.security/assessments' `
                                | mvexpand Category=properties.metadata.categories `
                                | extend AssessmentId=id, `
                                    AssessmentKey=name, `
                                    ResourceId=properties.resourceDetails.Id, `
                                    ResourceIdsplit = split(properties.resourceDetails.Id,'/'), `
	                                RecommendationId=name, `
	                                RecommendationName=properties.displayName, `
	                                Source=properties.resourceDetails.Source, `
	                                RecommendationState=properties.status.code, `
	                                ActionDescription=properties.metadata.description, `
	                                AssessmentType=properties.metadata.assessmentType, `
	                                RemediationDescription=properties.metadata.remediationDescription, `
	                                PolicyDefinitionId=properties.metadata.policyDefinitionId, `
	                                ImplementationEffort=properties.metadata.implementationEffort, `
	                                RecommendationSeverity=properties.metadata.severity, `
                                    Threats=properties.metadata.threats, `
	                                UserImpact=properties.metadata.userImpact, `
	                                AzPortalLink=properties.links.azurePortal, `
	                                MoreInfo=properties `
                                | extend ResourceSubId = tostring(ResourceIdsplit[(2)]), `
                                    ResourceRgName = tostring(ResourceIdsplit[(4)]), `
                                    ResourceType = tostring(ResourceIdsplit[(6)]), `
                                    ResourceName = tostring(ResourceIdsplit[(8)]), `
                                    FirstEvaluationDate = MoreInfo.status.firstEvaluationDate, `
                                    StatusChangeDate = MoreInfo.status.statusChangeDate, `
                                    Status = MoreInfo.status.code `
                                | join kind=leftouter (resourcecontainers | where type=='microsoft.resources/subscriptions' | project SubName=name, subscriptionId) on subscriptionId `
	                            | where AssessmentType == 'BuiltIn' `
                                | project-away kind,managedBy,sku,plan,tags,identity,zones,location,ResourceIdsplit,id,name,type,resourceGroup,subscriptionId, extendedLocation,subscriptionId1 `
                                | project SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,TenantId=tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink, AssessmentKey `
                                | where RecommendationState == 'Unhealthy' `
                                | join kind=leftouter (
	                                securityresources
	                                | where type == 'microsoft.security/assessments/subassessments'
	                                | extend AssessmentKey = extract('.*assessments/(.+?)/.*',1,  id)
                                        | project AssessmentKey, subassessmentKey=name, id, parse_json(properties), resourceGroup, subscriptionId, tenantId
                                        | extend SubAssessmentSescription = properties.description,
                                            SubAssessmentDisplayName = properties.displayName,
                                            SubAssessmentResourceId = properties.resourceDetails.id,
                                            SubAssessmentResourceSource = properties.resourceDetails.source,
                                            SubAssessmentCategory = properties.category,
                                            SubAssessmentSeverity = properties.status.severity,
                                            SubAssessmentCode = properties.status.code,
                                            SubAssessmentTimeGenerated = properties.timeGenerated,
                                            SubAssessmentRemediation = properties.remediation,
                                            SubAssessmentImpact = properties.impact,
                                            SubAssessmentVulnId = properties.id,
                                            SubAssessmentMoreInfo = properties.additionalData,
                                            SubAssessmentMoreInfoAssessedResourceType = properties.additionalData.assessedResourceType,
                                            SubAssessmentMoreInfoData = properties.additionalData.data
                                ) on AssessmentKey"
                            First = $pageSize
                        }

    $results = do {
        $iteration += 1
        Write-Verbose "Iteration #$iteration" -Verbose
        $pageResults = Search-AzGraph  @searchParams -ManagementGroup $Global:ManagementGroupScope
        $searchParams.Skip += $pageResults.Count
        $MDC_Recommendations += $pageResults
    } while ($pageResults.Count -eq $pageSize)
#>

################################################
# (A2) MDC | Recommendations with link
################################################

    $MDC_Recommendations = @()
    $pageSize = 1000
    $iteration = 0
    $searchParams = @{
                        Query = "SecurityResources `
                                | where type == 'microsoft.security/assessments' `
                                | mvexpand Category=properties.metadata.categories `
                                | extend AssessmentId=id, `
                                    AssessmentKey=name, `
                                    ResourceId=properties.resourceDetails.Id, `
                                    ResourceIdsplit = split(properties.resourceDetails.Id,'/'), `
	                                RecommendationId=name, `
	                                RecommendationName=properties.displayName, `
	                                Source=properties.resourceDetails.Source, `
	                                RecommendationState=properties.status.code, `
	                                ActionDescription=properties.metadata.description, `
	                                AssessmentType=properties.metadata.assessmentType, `
	                                RemediationDescription=properties.metadata.remediationDescription, `
	                                PolicyDefinitionId=properties.metadata.policyDefinitionId, `
	                                ImplementationEffort=properties.metadata.implementationEffort, `
	                                RecommendationSeverity=properties.metadata.severity, `
                                    Threats=properties.metadata.threats, `
	                                UserImpact=properties.metadata.userImpact, `
	                                AzPortalLink=properties.links.azurePortal, `
	                                MoreInfo=properties `
                                | extend ResourceSubId = tostring(ResourceIdsplit[(2)]), `
                                    ResourceRgName = tostring(ResourceIdsplit[(4)]), `
                                    ResourceType = tostring(ResourceIdsplit[(6)]), `
                                    ResourceName = tostring(ResourceIdsplit[(8)]), `
                                    FirstEvaluationDate = MoreInfo.status.firstEvaluationDate, `
                                    StatusChangeDate = MoreInfo.status.statusChangeDate, `
                                    Status = MoreInfo.status.code `
                                | join kind=leftouter (resourcecontainers | where type=='microsoft.resources/subscriptions' | project SubName=name, subscriptionId) on subscriptionId `
	                            | where AssessmentType == 'BuiltIn' `
                                | project-away kind,managedBy,sku,plan,tags,identity,zones,location,ResourceIdsplit,id,name,type,resourceGroup,subscriptionId, extendedLocation,subscriptionId1 `
                                | project SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,TenantId=tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink, AssessmentKey `
                                | where RecommendationState == 'Unhealthy' "
                            First = $pageSize
                        }

    $results = do {
        $iteration += 1
        Write-Verbose "Iteration #$iteration" -Verbose
        $pageResults = Search-AzGraph  @searchParams -ManagementGroup $Global:ManagementGroupScope
        $searchParams.Skip += $pageResults.Count
        $MDC_Recommendations += $pageResults
    } while ($pageResults.Count -eq $pageSize)


#########################################################
# (A3) MDC | SubAssessments (Detailed Infomation)
#########################################################

    $MDC_Recommendations_SubAssessments = @()
    $pageSize = 1000
    $iteration = 0
    $searchParams = @{
                        Query = "SecurityResources `
	                            | where type == 'microsoft.security/assessments/subassessments'
	                            | extend AssessmentKey = extract('.*assessments/(.+?)/.*',1,  id)
                                | project AssessmentKey, subassessmentKey=name, id, parse_json(properties), resourceGroup, subscriptionId, tenantId
                                | extend SubAssessDescription = properties.description,
                                        SubAssessDisplayName = properties.displayName,
                                        SubAssessResourceId = properties.resourceDetails.id,
                                        SubAssessResourceSource = properties.resourceDetails.source,
                                        SubAssessCategory = properties.category,
                                        SubAssessSeverity = properties.status.severity,
                                        SubAssessCode = properties.status.code,
                                        SubAssessTimeGenerated = properties.timeGenerated,
                                        SubAssessRemediation = properties.remediation,
                                        SubAssessImpact = properties.impact,
                                        SubAssessVulnId = properties.id,
                                        SubAssessMoreInfo = properties.additionalData,
                                        SubAssessMoreInfoAssessedResourceType = properties.additionalData.assessedResourceType,
                                        SubAssessMoreInfoData = properties.additionalData.data `
                                | join kind=leftouter (resourcecontainers | where type=='microsoft.resources/subscriptions' | project SubName=name, subscriptionId) on subscriptionId "
                            First = $pageSize
                        }

    $results = do {
        $iteration += 1
        Write-Verbose "Iteration #$iteration" -Verbose
        $pageResults = Search-AzGraph  @searchParams -ManagementGroup $Global:ManagementGroupScope
        $searchParams.Skip += $pageResults.Count
        $MDC_Recommendations_SubAssessments += $pageResults
    } while ($pageResults.Count -eq $pageSize)

    ###################################################
    # (A3.1) SubAssessment Identity Lookup
    ###################################################

        # Identities
        $Identities_SubAssessments_Full  = $MDC_Recommendations_SubAssessments | where-Object { $_.SubAssessResourceId -like "*/identities/*" }

        $Identities_SubAssessments_Obj = @()
        ForEach ($Obj in $Identities_SubAssessments_Full)
            {
                $Identities_SubAssessments_Obj  += ($Obj.SubAssessResourceId.split("/"))[4]
            }
        $Identities_SubAssessments_List = $Identities_SubAssessments_Obj | Sort-Object -Unique

        $IdArray = @()
        ForEach ($IdObj in $Identities_SubAssessments_List)
            {
                $ObjInfo = Get-AzureADObjectByObjectId -ObjectIds $IdObj
                $Object = New-Object -TypeName PSObject
                $Object | Add-Member -MemberType NoteProperty -Name ObjectId -value $ObjInfo.ObjectId
                $Object | Add-Member -MemberType NoteProperty -Name AccountEnabled -value $ObjInfo.AccountEnabled
                $Object | Add-Member -MemberType NoteProperty -Name UserPrincipalName -value $ObjInfo.UserPrincipalName
                $Object | Add-Member -MemberType NoteProperty -Name DisplayName -value $ObjInfo.DisplayName
                Write-Output "Checking $($ObjInfo.UserPrincipalName)"
                $IdArray += $Object
            }

        $MDC_Recommendations_SubAssessments_Count = $MDC_Recommendations_SubAssessments.count
        $Iteration = 0

        Do 
            {
                Write-Output "Processing $($Iteration) / $($MDC_Recommendations_SubAssessments_Count) ..."
                $Iteration += 1
                If ($MDC_Recommendations_SubAssessments[$Iteration].SubAssessResourceId -like "*/identities/*")
                    {
                        $Identities_SubAssessments_Obj  = $MDC_Recommendations_SubAssessments[$Iteration].SubAssessResourceId.split("/")[4]
                        $UserInfo = $IdArray | Where-Object { $_.ObjectId -eq $Identities_SubAssessments_Obj } 
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResObjectId -value $UserInfo.ObjectId -Force
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResAccountEnabled -value $UserInfo.AccountEnabled -Force
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResUserPrincipalName -value $UserInfo.UserPrincipalName -Force
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResDisplayName -value $UserInfo.DisplayName -Force
                    }
                Else
                    {
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResObjectId -value "" -Force
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResAccountEnabled -value "" -Force
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResUserPrincipalName -value "" -Force
                        $MDC_Recommendations_SubAssessments[$Iteration] | Add-Member -MemberType NoteProperty -Name SubAssessResDisplayName -value "" -Force
                    }
            }
        Until ($Iteration -eq $MDC_Recommendations_SubAssessments_Count)


#######################
# (B1) Filtering of Scope
#######################

    $MDC_Recommendations_Unhealthy        = $MDC_Recommendations | Where-Object {$_.MoreInfo.status.code -eq "Unhealthy" }

    # Scope for pivot
    $Category_SubLevel                    = $MDC_Recommendations_Unhealthy | Where-Object { ($_.resourceName -eq $null) -or ($_.resourceName -eq "") } | Select-Object SubName, ResourceSubId, tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_ResourceLevel               = $MDC_Recommendations_Unhealthy | Where-Object { ($_.resourceName -ne $null) -and ($_.resourceName -ne "") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_Networking                  = $MDC_Recommendations_Unhealthy | Where-Object { ($_.Category -eq "Networking") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_AppServices                 = $MDC_Recommendations_Unhealthy | Where-Object { ($_.Category -eq "AppServices") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_Compute                     = $MDC_Recommendations_Unhealthy | Where-Object { ($_.Category -eq "Compute") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_Container                   = $MDC_Recommendations_Unhealthy | Where-Object { ($_.Category -eq "Container") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_Data                        = $MDC_Recommendations_Unhealthy | Where-Object { ($_.Category -eq "Data") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_IoT                         = $MDC_Recommendations_Unhealthy | Where-Object { ($_.Category -eq "IoT") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_IdentityAndAccess           = $MDC_Recommendations_Unhealthy | Where-Object { ($_.Category -eq "IdentityAndAccess") } | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink 
    $Category_Other                       = $MDC_Recommendations_Unhealthy | Where-Object { ( ($_.Category -ne "Networking") -and `
                                                                                              ($_.Category -ne "AppServices") -and `
                                                                                              ($_.Category -ne "Compute") -and `
                                                                                              ($_.Category -ne "Container") -and `
                                                                                              ($_.Category -ne "Data") -and `
                                                                                              ($_.Category -ne "IoT") -and `
                                                                                              ($_.Category -ne "IdentityAndAccess") ) }  | Select-Object SubName, ResourceSubId, ResourceRgName,ResourceType,ResourceName,tenantId, RecommendationName, RecommendationId, RecommendationState, RecommendationSeverity, AssessmentType, PolicyDefinitionId, ImplementationEffort, UserImpact, Category, Threats, Source, ActionDescription, RemediationDescription, MoreInfo, ResourceId, AzPortalLink
                                                                                              


    $RBAC_Scope_Delegation_Direct         = $RBAC_RoleAssignments | Where-Object { $_.Scope_Delegation -like "Direct*" }
    $RBAC_Scope_Delegation_Inheritance_MG = $RBAC_RoleAssignments | Where-Object { $_.Scope_Delegation -eq "Inheritance_MG" }

    $RBAC_Delegation_Type_Direct_Sub      = $RBAC_RoleAssignments | Where-Object { ($_.RBAC_Delegation_Type -eq "Direct") -and ($_.Scope_Delegation -eq 'Direct_SUB') }
    
    # Sub - RBAC_Delegation_Type_Direct_FilterAway

        $RBAC_Delegation_Type_Direct_Sub_Filtered = $RBAC_Delegation_Type_Direct_Sub

<#
        Enable these 2 filters, if you want to remove these from the overview
            
        # Filter away User Access Administrator Role Definininition Name (inherited)
        $RBAC_Delegation_Type_Direct_Sub_Filtered = $RBAC_Delegation_Type_Direct_Sub_Filtered | Where-Object { $_.RoleDefinitionName -ne "User Access Administrator" }

        # Filter away DisplayName startswith Defender*
        $RBAC_Delegation_Type_Direct_Sub_Filtered = $RBAC_Delegation_Type_Direct_Sub_Filtered | Where-Object { $_.DisplayName -notlike "Defender*" }
#>

    $RBAC_Delegation_Type_Direct_Mg      = $RBAC_RoleAssignments | Where-Object { ($_.RBAC_Delegation_Type -eq "Direct") -and ($_.Scope_Delegation -eq 'Inheritance_MG') }
    
        # MG - RBAC_Delegation_Type_Direct_FilterAway

        $RBAC_Delegation_Type_Direct_Mg_Filtered = $RBAC_Delegation_Type_Direct_Mg
    
<#
        Enable this filter, if you want to remove these from the overview

        # Filter away DisplayName startswith Defender*
        $RBAC_Delegation_Type_Direct_Mg_Filtered = $RBAC_Delegation_Type_Direct_Mg_Filtered | Where-Object { $_.DisplayName -notlike "Defender*" }
#>

#########################################################
# (C1) Reporting | Export to Excel
#########################################################

    Remove-Item $FileOutput

    #--------------------------------------

    # Content
    # Purpose: Show introduction of tables and pivots

        $TableName = "Introduction_Help"
        $TableSelection = Import-csv $HelpContentFile -Delimiter ";" -Encoding UTF8
        Write-Output "Exporting $($TableName) .... Please Wait !"
        $excel = $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13 -PassThru
        $excel.Workbook.Worksheets[1].Column(1) | Set-Format -VerticalAlignment Top -HorizontalAlignment Left -WrapText -Width 35
        $excel.Workbook.Worksheets[1].Column(2) | Set-Format -VerticalAlignment Top -HorizontalAlignment Left -WrapText -Width 90
        $excel.Workbook.Worksheets[1].Column(3) | Set-Format -VerticalAlignment Top -HorizontalAlignment Left -WrapText -Width 40
        Close-ExcelPackage $excel


    #--------------------------------------

    # Unhealthy_All
    # Purpose: Show All Unhealthy recommendations

        $TableName = "Unhealthy_All"
        $TableSelection = $MDC_Recommendations_Unhealthy
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }
    #--------------------------------------

    # Unhealthy_High
    # Purpose: Show all Unhealthy recommendations with High priority

        $TableName = "Unhealthy_High"
        $TableSelection = $MDC_Recommendations_Unhealthy | Where-Object { $_.recommendationSeverity -eq "High" }
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------

    # Unhealthy_Medium
    # Purpose: Show all Unhealthy recommendations with medium priority

        $TableName = "Unhealthy_Medium"
        $TableSelection = $MDC_Recommendations_Unhealthy | Where-Object { $_.recommendationSeverity -eq "Medium" }
        Write-Output "Exporting $($TableName) .... Please Wait !"
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------

    # Unhealthy_Low
    # Purpose: Show all Unhealthy recommendations with Low priority

        $TableName = "Unhealthy_Low"
        $TableSelection = $MDC_Recommendations_Unhealthy | Where-Object { $_.recommendationSeverity -eq "Low" }
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------

    # Unhealthy_All_SubLevel
    # Purpose: Show all Unhealthy recommendations where the target is on subcription-level

        $TableName = "Unhealthy_All_SubLevel"
        $TableSelection = $Category_SubLevel
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_All_ResourceLevel
    # Purpose: Show all Unhealthy recommendations where the target is on resource-level

        $TableName = "Unhealthy_All_ResourceLevel"
        $TableSelection = $Category_ResourceLevel
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_Networking
    # Purpose: Show all Unhealthy recommendations related to Networking category

        $TableName = "Unhealthy_Category_Networking"
        $TableSelection = $Category_Networking
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_AppServices
    # Purpose: Show all Unhealthy recommendations related to AppServices category

        $TableName = "Unhealthy_Category_AppServices"
        $TableSelection = $Category_AppServices
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_Compute
    # Purpose: Show all Unhealthy recommendations related to Compute category

        $TableName = "Unhealthy_Category_Compute"
        $TableSelection = $Category_Compute
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_Container
    # Purpose: Show all Unhealthy recommendations related to Container category

        $TableName = "Unhealthy_Category_Container"
        $TableSelection = $Category_Container
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_Data
    # Purpose: Show all Unhealthy recommendations related to Data category

        $TableName = "Unhealthy_Category_Data"
        $TableSelection = $Category_Data
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_IoT
    # Purpose: Show all Unhealthy recommendations related to IoT category

        $TableName = "Unhealthy_Category_IoT"
        $TableSelection = $Category_IoT
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_Id_Access
    # Purpose: Show all Unhealthy recommendations related to IdentityAndAccess category

        $TableName = "Unhealthy_Category_Id_Access"
        $TableSelection = $Category_IdentityAndAccess
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Unhealthy_Category_Other
    # Purpose: Show all Unhealthy recommendations related to Other categories

        $TableName = "Unhealthy_Category_Other"
        $TableSelection = $Category_Other
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # SubAssess_All
    # Purpose: Show all detailed information (SubAssessments)

        $TableName = "SubAssess_All"
        $TableSelection = $MDC_Recommendations_SubAssessments | Select-Object SubName,subscriptionId,SubAssessDisplayName,SubAssessResourceId,SubAssessResObjectId,SubAssessResDisplayName,SubAssessResUserPrincipalName,SubAssessResAccountEnabled,SubAssessResourceSource,SubAssessCategory,SubAssessSeverity,SubAssessCode,SubAssessTimeGenerated,SubAssessRemediation,SubAssessImpact,SubAssessVulnId
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # SubAssess_Identity
    # Purpose: Show all identity detailed information (SubAssessments).

        $TableName = "SubAssess_Identity"
        $TableSelection = $MDC_Recommendations_SubAssessments | Where-Object { $_.SubAssessResUserPrincipalName -ne "" } | Select-Object SubName,subscriptionId,SubAssessDisplayName,SubAssessResourceId,SubAssessResObjectId,SubAssessResDisplayName,SubAssessResUserPrincipalName,SubAssessResAccountEnabled,SubAssessResourceSource,SubAssessCategory,SubAssessSeverity,SubAssessCode,SubAssessTimeGenerated,SubAssessRemediation,SubAssessImpact,SubAssessVulnId
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # SubAssess_Other
    # Purpose: Show all other, non-identity detailed information (SubAssessments)

        $TableName = "SubAssess_Other"
        $TableSelection = $MDC_Recommendations_SubAssessments | Where-Object { $_.SubAssessResUserPrincipalName -eq "" } | Select-Object SubName,subscriptionId,SubAssessDisplayName,SubAssessResourceId,SubAssessResObjectId,SubAssessResDisplayName,SubAssessResUserPrincipalName,SubAssessResAccountEnabled,SubAssessResourceSource,SubAssessCategory,SubAssessSeverity,SubAssessCode,SubAssessTimeGenerated,SubAssessRemediation,SubAssessImpact,SubAssessVulnId
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # RBAC_RoleAssignments
    # Purpose: Show all Azure RBAC Role Assignments

        $TableName = "RBAC_RoleAssignments"
        $TableSelection = $RBAC_RoleAssignments
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }


    #--------------------------------------
    # RBAC_Direct_Sublevel
    # Purpose: Show all Azure RBAC Role Assignments directly on Sub-level (not part of group)

        $TableName = "RBAC_Direct_Sublevel"
        $TableSelection = $RBAC_Delegation_Type_Direct_Sub_Filtered
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # RBAC_Direct_Mglevel
    # Purpose: Show all Azure RBAC Role Assignments directly on Mg-level (not part of group)

        $TableName = "RBAC_Direct_Mglevel"
        $TableSelection = $RBAC_Delegation_Type_Direct_Mg_Filtered
        If ($TableSelection)
            {
                Write-Output "Exporting $($TableName) .... Please Wait !"
                $TableSelection | Export-Excel -Path $FileOutput -WorksheetName $TableName -AutoFilter -AutoSize -BoldTopRow -tablename $TableName -tablestyle Medium13
            }

    #--------------------------------------
    # Excel export - prepare pivots

        Write-Output "Preparing Pivot tables .... Please Wait !"
        $Pivottable = [ordered]@{}

    #--------------------------------------

    <#
        PT_CATEGORY_SUBLEVEL
          Purpose: Prioritize recommendations based on Category and RecommendationSeverity
          SourceWorkSheet: Unhealthy_All_SubLevel
          Sort-order:
                   (1) Category (Storage, Network, Identity, etc.)
                   (2) RecommendationSeverity (High, Medium, Low)
                   (3) RecommendationName
                   (4) SubName (suscription name)
    #>
        If ($Category_SubLevel)
            {
                $Pivottable.PT_CATEGORY_SUBLEVEL      = @{ SourceWorkSheet = "Unhealthy_All_SubLevel"
                                                           PivotRows       = @('Category','RecommendationSeverity','RecommendationName','SubName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }


    #--------------------------------------

    <#
        PT_CATEGORY_RESOURCELEVEL
          Purpose: Prioritize recommendations based on ResourceType and RecommendationSeverity
          SourceWorkSheet: Unhealthy_All_ResourceLevel
          Sort-order:
                   (1) Category (Storage, Network, Identity, etc.)
                   (2) RecommendationSeverity (High, Medium, Low)
                   (3) RecommendationName
                   (4) SubName (suscription name)
                   (5) ResourceRgName
                   (6) ResourceName
    #>

        If ($Category_ResourceLevel)
            {
                $Pivottable.PT_CATEGORY_RESOURCELEVEL = @{ SourceWorkSheet = "Unhealthy_ALL_ResourceLevel"
                                                           PivotRows       = @('Category','RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_CATEGORY_NETWORKING
          Purpose: Prioritize networking recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_Networking
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_Networking)
            {
                $Pivottable.PT_CATEGORY_NETWORKING    = @{ SourceWorkSheet = "Unhealthy_Category_Networking"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_CATEGORY_APPSERVICES
          Purpose: Prioritize AppServices recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_AppServices
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_AppServices)
            {
                $Pivottable.PT_CATEGORY_APPSERVICES   = @{ SourceWorkSheet = "Unhealthy_Category_AppServices"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_CATEGORY_COMPUTE
          Purpose: Prioritize Compute recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_Compute
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_Compute)
            {
                $Pivottable.PT_CATEGORY_COMPUTE       = @{ SourceWorkSheet = "Unhealthy_Category_Compute"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }

            }

    #--------------------------------------

    <#
        PT_CATEGORY_CONTAINER
          Purpose: Prioritize Container recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_Container
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_Container)
            {
                $Pivottable.PT_CATEGORY_CONTAINER     = @{ SourceWorkSheet = "Unhealthy_Category_Container"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_CATEGORY_DATA
          Purpose: Prioritize Data recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_Data
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_Data)
            {
                $Pivottable.PT_CATEGORY_DATA          = @{ SourceWorkSheet = "Unhealthy_Category_Data"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_CATEGORY_ID_ACCESS
          Purpose: Prioritize Identity and Access recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_IdentityAndAccess
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_IdentityAndAccess)
            {
                $Pivottable.PT_CATEGORY_ID_ACCESS     = @{ SourceWorkSheet = "Unhealthy_Category_Id_Access"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_CATEGORY_IOT
          Purpose: Prioritize IoT recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_IoT
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_IoT)
            {
                $Pivottable.PT_CATEGORY_IOT           = @{ SourceWorkSheet = "Unhealthy_Category_IoT"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_CATEGORY_OTHER
          Purpose: Prioritize other recommendations based on RecommendationSeverity
          SourceWorkSheet: Unhealthy_Category_Other
          Sort-order:
                   (1) RecommendationSeverity (High, Medium, Low)
                   (2) RecommendationName
                   (3) SubName (suscription name)
                   (4) ResourceRgName
                   (5) ResourceName
    #>

        If ($Category_Other)
            {
                $Pivottable.PT_CATEGORY_OTHER         = @{ SourceWorkSheet = "Unhealthy_Category_Other"
                                                           PivotRows       = @('RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_RESOURCETYPE
          Purpose: Prioritize recommendations based on ResourceType and RecommendationSeverity
          SourceWorkSheet: Unhealthy_All_ResourceLevel
          Sort-order:
                   (1) ResourceType (SQL, Keyvault, Storage, Network, Identity, etc.)
                   (2) RecommendationSeverity (High, Medium, Low)
                   (3) RecommendationName
                   (4) SubName (suscription name)
                   (5) ResourceRgName
                   (6) ResourceName
    #>

        If ($Category_ResourceLevel)
            {
                $Pivottable.PT_RESOURCETYPE           = @{ SourceWorkSheet = "Unhealthy_All_ResourceLevel"
                                                           PivotRows       = @('ResourceType','RecommendationSeverity','RecommendationName','SubName','ResourceRgName','ResourceName')
                                                           PivotData       = @{"RecommendationId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_SUBASSESS_IDENTITY
          Purpose: Prioritize recommendations based on SubAssessmentSeverity (identity-related)
          SourceWorkSheet: Unhealthy_All_ResourceLevel
          Sort-order:
                   (1) SubAssessSeverity (High, Medium, Low)
                   (2) SubAssessDisplayName (recommendation)
                   (3) SubAssessResDisplayName (resource)
                   (4) SubAssessResUserPrincipalName (resource)
                   (5) SubName (suscription name)
    #>

        If ($MDC_Recommendations_SubAssessments)
            {
                $Pivottable.PT_SUBASSESS_IDENTITY     = @{ SourceWorkSheet = "SubAssess_Identity"
                                                           PivotRows       = @('SubAssessSeverity','SubAssessDisplayName','SubAssessResDisplayName','SubAssessResUserPrincipalName','SubName')
                                                           PivotData       = @{"SubAssessCode"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_SUBASSESS_OTHER
          Purpose: Prioritize recommendations based on SubAssessmentSeverity (non-identity related)
          SourceWorkSheet: SubAssess_Other
          Sort-order:
                   (1) SubAssessSeverity (High, Medium, Low)
                   (2) SubAssessDisplayName (recommendation)
                   (3) SubAssessResDisplayName (resource)
                   (4) SubAssessResUserPrincipalName (resource)
                   (5) SubName (suscription name)
    #>

        If ($MDC_Recommendations_SubAssessments)
            {
                $Pivottable.PT_SUBASSESS_OTHER        = @{ SourceWorkSheet = "SubAssess_Other"
                                                           PivotRows       = @('SubAssessSeverity','SubAssessDisplayName','SubName','SubAssessResourceId')
                                                           PivotData       = @{"SubAssessCode"="Count"}
                                                         }
            }

    #--------------------------------------
    <#
        PT_RBAC_ROLEDEF
          Purpose: Detail RBAC permissions, sorted by RoleDefinitionName and Scope_Delegation
          SourceWorkSheet: RBAC_RoleAssignments
          Sort-order:
                   (1) SubscriptionName
                   (2) RoleDefinitionName (Contributor, Owner, Cost Management Reader, etc.)
                   (3) Scope_Delegation (Direct_SUB, Direct_RG, Inheritance_MG)
                   (4) Scope (management group name or /)
                   (5) RBAC_Delegation_Type (Group_inheritance, Direct)
                   (6) RBAC_GroupName
                   (7) ObjectType
                   (8) DisplayName
                   (9) UserPrincipalName
    #>

        If ($RBAC_RoleAssignments)
            {
                $Pivottable.PT_RBAC_ROLEDEF           = @{ SourceWorkSheet = "RBAC_RoleAssignments"
                                                           PivotRows       = @('SubscriptionName','RoleDefinitionName','Scope_Delegation','Scope','RBAC_Delegation_Type','RBAC_GroupName','ObjectType','DisplayName','UserPrincipalName')
                                                           PivotData       = @{"ObjectId"="Count"}
                                                         }
            }

    #--------------------------------------
    <#
        PT_RBAC_SCOPE_DELEGATION
          Purpose: Detail RBAC permissions, sorted by Scope_Delegation and RoleDefinitionName
          SourceWorkSheet: RBAC_RoleAssignments
          Sort-order:
                   (1) SubscriptionName
                   (2) Scope_Delegation (Direct_SUB, Direct_RG, Inheritance_MG)
                   (3) RoleDefinitionName (Contributor, Owner, Cost Management Reader, etc.)
                   (4) Scope (management group name or /)
                   (5) RBAC_Delegation_Type (Group_inheritance, Direct)
                   (6) RBAC_GroupName
                   (7) ObjectType
                   (8) DisplayName
                   (9) UserPrincipalName
    #>

        If ($RBAC_RoleAssignments)
            {
                $Pivottable.PT_RBAC_SCOPE_DELEGATION  = @{ SourceWorkSheet = "RBAC_RoleAssignments"
                                                           PivotRows       = @('SubscriptionName','Scope_Delegation','RoleDefinitionName','Scope','RBAC_Delegation_Type','RBAC_GroupName','ObjectType','DisplayName','UserPrincipalName')
                                                           PivotData       = @{"ObjectId"="Count"}
                                                         }
            }

    #--------------------------------------

    <#
        PT_RBAC_MG
          Purpose: Detect direct RBAC permissions on Mg-level (not done by group)
          SourceWorkSheet: RBAC_Direct_Mglevel
          Sort-order:
                   (1) SubscriptionName
                   (2) RoleDefinitionName (Contributor, Owner, Cost Management Reader, etc.)
                   (3) Scope_Delegation (Direct_SUB, Direct_RG, Inheritance_MG)
                   (4) Scope (management group name or /)
                   (5) ObjectType
                   (6) DisplayName
                   (7) UserPrincipalName
    #>

        If ($RBAC_Delegation_Type_Direct_Mg_Filtered)
            {
                $Pivottable.PT_RBAC_DIRECT_MG         = @{ SourceWorkSheet = "RBAC_Direct_Mglevel"
                                                           PivotRows       = @('SubscriptionName','RoleDefinitionName','Scope_Delegation','Scope','ObjectType','DisplayName','UserPrincipalName')
                                                           PivotData       = @{"ObjectId"="Count"}
                                                         }

            }

    #--------------------------------------

    <#
        PT_RBAC_SUB
          Purpose: Detect direct RBAC permissions on Sub-level (not done by group)
          SourceWorkSheet: RBAC_Direct_Sublevel
          Sort-order:
                   (1) SubscriptionName
                   (2) RoleDefinitionName (Contributor, Owner, Cost Management Reader, etc.)
                   (3) Scope_Delegation (Direct_SUB, Direct_RG, Inheritance_MG)
                   (4) Scope (management group name or /)
                   (5) ObjectType
                   (6) DisplayName
                   (7) UserPrincipalName
    #>

        If ($RBAC_Delegation_Type_Direct_Sub_Filtered)
            {
                $Pivottable.PT_RBAC_DIRECT_SUB        = @{ SourceWorkSheet = "RBAC_Direct_Sublevel"
                                                           PivotRows       = @('SubscriptionName','RoleDefinitionName','Scope_Delegation','Scope','ObjectType','DisplayName','UserPrincipalName')
                                                           PivotData       = @{"ObjectId"="Count"}
                                                         }
            }

    #--------------------------------------
    # Exporting Excel pivot tables

        Write-Output "Exporting Pivot tables .... Please Wait !"
        Export-Excel -Path $FileOutput -PivotTableDefinition $Pivottable -NoTotalsInPivot

    # Finished
        Write-Output ""
        Write-Output "Export finished - $($FileOutput)"
