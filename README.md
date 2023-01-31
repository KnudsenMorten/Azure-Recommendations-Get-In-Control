# Azure-Recommendations-Get-In-Control
 Automate Reporting of Defender for Cloud recommendations & Role Assignments with 35 different views
 
    .SYNOPSIS
    
    In this script, I will demonstrate, how you can extract security recommendations from Microsoft
    Defender for Cloud using Azure Resource Graph - delivering a horizontal cross-subscriptions, 
    workload overview. Data will automatically be exported into a Excel spreadsheet delivering 19 Excel
    tables and 16 pivot tables.

    Information can be used to detect deviations from best practice / desired state - covering
    * Getting-in-control with workloads in tenant/management group (storage, network, app services, 
    containers, etc.) where we are not in control according to security best practice / desired state
    * Getting-in-control with subscriptions, where environment are not configured according to security
    best practice / desired state
    * Get-in-control with role assignments in tenant / management group / subscription / resource group.
    * Get detailed information about role assignments on user / service-principal-level, based on direct
    assignment and inheritance
    * Get detailed insight about users / service-principal-level, based on group membership - both direct
    and inheritance.

    .NOTES
    VERSION: 2301

    .COPYRIGHT
    @mortenknudsendk on Twitter
    Blog: https://mortenknudsen.net
    
    .LICENSE
    Licensed under the MIT license.

    .WARRANTY
    Use at your own risk, no warranty given!
