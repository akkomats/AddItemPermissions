<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
Param(
    $SiteUrl,
    $LibraryName,
    $ItemID,
    $UserName,
    $RoleName
)

If ($SiteUrl -eq $null)
{
   Write-Host "Example)"
   Write-Host ">.\SetItemPermissions.ps1 -siteUrl 'https://tenant.sharepoint.com/sites/site' -LibraryName 'DocumentLibrary1' -ItemID '2' -UserName 'user1@tenant.onmicrosoft.com' -RoleName 'Edit'"
   return
}

#---------------------------------------
# Initialize
#---------------------------------------
$ErrorActionPreference = "Stop";
$clientAssembly = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client");
$clientRuntimeAssembly = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime");
$assemblies = ($clientAssembly.FullName, $clientRuntimeAssembly.FullName);

# Helper Method
Add-Type -Language CSharp -ReferencedAssemblies $assemblies -TypeDefinition "
using Microsoft.SharePoint.Client;
public static class Helper
{
    public static void LoadHasUniqueRoleAssignments(ClientContext context, SecurableObject securableObject)
    {
        context.Load(securableObject,
            o => o.HasUniqueRoleAssignments);
    }
}";

function ExecuteQueryWithIncrementalRetry($retryCount, $delay)
{
  $retryAttempts = 0;
  $backoffInterval = $delay;
  if ($retryCount -le 0)
  {
    throw "Provide a retry count greater than zero."
  }
  if ($delay -le 0)
  {
    throw "Provide a delay greater than zero."
  }
  while ($retryAttempts -lt $retryCount)
  {
    try
    {
      $script:context.ExecuteQuery()
      return;
    }
    catch [System.Net.WebException]
    {
      $response = $_.Exception.Response
      if ($response -ne $null -and $response.StatusCode -eq 429)
      {
        Write-Host ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($backoffInterval/1000))
        #Add delay.
        Start-Sleep -m $backoffInterval
        #Add to retry count and increase delay.
        $retryAttempts++;
        $backoffInterval = $backoffInterval * 2;
      }
      else
      {
        throw;
      }
    }
  }
  throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}

#---------------------------------------
# Main
#---------------------------------------
$context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl);

try
{
    Write-Host "Please input user name : ";
    $exeUserName = Read-Host;

    Write-Host "Please input password : ";
    $password = Read-Host -AsSecureString;

    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($exeUserName, $password);
    $context.Credentials = $credentials;

    $web = $context.Web
    $context.Load($web)

    $list = $web.Lists.GetByTitle($LibraryName);
    $context.Load($list);

    $roleDef = $context.Web.RoleDefinitions.GetByName($RoleName);
    $context.Load($roleDef);

    $user = $context.Web.EnsureUser($userName);
    $context.Load($user);

    $exeUser = $context.Web.EnsureUser($exeUserName);
    $context.Load($exeUser);

    $item = $list.GetItemByID($ItemID);
    $context.Load($item);

    ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000;

    # Check the item has unique role assignments.
    [Helper]::LoadHasUniqueRoleAssignments($context, $item);
    ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000;
        
    # If the item does not have unique role assignments, delete role inheritance.
    if($item.HasUniqueRoleAssignments -eq $False)
    {   
        $item.BreakRoleInheritance($false, $false);
        $collRoleDefinitionBinding = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context);
        $collRoleDefinitionBinding.Add($roleDef);
        $roleAssignment = $item.RoleAssignments.Add($user, $collRoleDefinitionBinding);

        $item.RoleAssignments.GetByPrincipal($exeUser).DeleteObject();        
        ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000;
    }
    else
    {
        # If the item has unique role assignments, add role.
        $collRoleDefinitionBinding = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context);
        $collRoleDefinitionBinding.Add($roleDef);
        $roleAssignment = $item.RoleAssignments.Add($user, $collRoleDefinitionBinding);
        ExecuteQueryWithIncrementalRetry -retryCount 5 -delay 30000;
    }

    $context.Dispose();
}
catch
{
    Write-Host "$($error[0].Exception)`r`n$($Error[0].InvocationInfo.PositionMessage)" -ForegroundColor Red;
    $response = $_.Exception.Response;
    $response.Headers;
}
finally
{
    if($context -ne $null)
    {
        $context.Dispose();
    }
}
