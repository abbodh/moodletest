<#
    File Name :  Moodle-AzureAD-Script.ps1
    
    Copyright (c) Microsoft Corporation. All rights reserved.
    Licensed under the MIT License.
#>


# Allow for the script to be run
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

# Install necessary modules
Install-Module AzureAD -AllowClobber -Scope CurrentUser
Install-Module Az -AllowClobber -Scope CurrentUser

#Overarching requirement - log into Azure first!
Connect-AzureAD

<#
.DESCRIPTION 

This function will be able to create an array of type RequiredResourceAccess which will be then passed to the New-AzureADApplication cmdlet
#>
function Get-Resources
{
    [Microsoft.Open.AzureAD.Model.RequiredResourceAccess[]] $outputArray = @();
    
    $localPath = Get-Location
    $jsonPath = -Join($localPath,'\Json\permissions.json');
    $jsonObj = (New-Object System.Net.WebClient).DownloadString($jsonPath) | ConvertFrom-Json;

    # Output the number of objects to push into the array outputArray
    Write-Host 'From the json path:'$jsonPath', we can find' $jsonObj.requiredResourceAccess.length'attributes to populate';

    for ($i = 0; $i -lt $jsonObj.requiredResourceAccess.length; $i++) {
        
        # Step A - Create a new object fo the type RequiredResourceAccess
        $reqResourceAccess = New-Object -TypeName Microsoft.Open.AzureAD.Model.RequiredResourceAccess; 

        # Step B - Straightforward setting the ResourceAppId accordingly
        $reqResourceAccess.ResourceAppId = $jsonObj.requiredResourceAccess[$i].resourceAppId;

        # Step C - Having to set the ResourceAccess carefully
        if ($jsonObj.requiredResourceAccess[$i].resourceAccess.length -gt 1)
        {
            $reqResourceAccess.ResourceAccess = $jsonObj.requiredResourceAccess[$i].resourceAccess;
        }
        else
        {
            $reqResourceAccess.ResourceAccess = $jsonObj.requiredResourceAccess[$i].resourceAccess[0];
        }

        # Step D - Add the element to the array
        $outputArray += $reqResourceAccess;
    }

    $outputArray;
}

<#
.DESCRIPTION 

This function will allow to create and add Microsoft Graph scope 
#>
function Create-Scope(
        [string] $value,
        [string] $userConsentDisplayName,
        [string] $userConsentDescription,
        [string] $adminConsentDisplayName,
        [string] $adminConsentDescription) {
        $scope = New-Object Microsoft.Open.MsGraph.Model.PermissionScope
        $scope.Id = New-Guid
        $scope.Value = $value
        $scope.UserConsentDisplayName = $userConsentDisplayName
        $scope.UserConsentDescription = $userConsentDescription
        $scope.AdminConsentDisplayName = $adminConsentDisplayName
        $scope.AdminConsentDescription = $adminConsentDescription
        $scope.IsEnabled = $true
        $scope.Type = "User"
        return $scope
    }

<#
.DESCRIPTION 

This function will allow to add preauthroized application for Azure AD application 
#>
function Create-PreAuthorizedApplication(
        [string] $applicationIdToPreAuthorize,
        [string] $scopeId) {
        $preAuthorizedApplication = New-Object 'Microsoft.Open.MSGraph.Model.PreAuthorizedApplication'
        $preAuthorizedApplication.AppId = $applicationIdToPreAuthorize
        $preAuthorizedApplication.DelegatedPermissionIds = @($scopeId)
        return $preAuthorizedApplication
    }

<#
.DESCRIPTION 

This function will check whether the URL is valid and secure 
#>
function IsValidateSecureUrl {
    param(
        [Parameter(Mandatory = $true)] [string] $url
    )
    # Url with https prefix REGEX matching
    return ($url -match "https:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)")
}

# Step 1 - Getting the necessary information
$displayName = Read-Host -Prompt "Enter the AAD app name (ex: Moodle plugin)"
$moodleDomain = Read-Host -Prompt "Enter the URL of your Moodle server (ex: https://www.moodleserver.com)"

if (-not(IsValidateSecureUrl($moodleUrl))) {
        Write-Host "Invalid websiteUrl. This should be an https url."
}

if ($moodleDomain -notmatch '.+?\/$')
{
    $moodleDomain += '/'
}

# Step 2 - Construct the reply URLs
$botFrameworkUrl = 'https://token.botframework.com/.auth/web/redirect'
$authUrl = $moodleUrl + '/auth/oidc/'


$replyUrls = ($moodleUrl, $botFrameworkUrl, $authUrl)

# Step 3 - Compile the Required Resource Access object
[Microsoft.Open.AzureAD.Model.RequiredResourceAccess[]] $requiredResourceAccess = Get-Resources

# Step 4 - Making sure to officially register the application
$app = New-AzureADApplication -DisplayName $displayName 

# Set identifier URL
$appId = $app.AppId
$appObjectId = $app.ObjectId
$IdentifierUris = 'api://' + $moodleDomain + '/' + $appId

# Removing default scope user_impersonation, access and optional claims
$localPath = Get-Location
$resetOptionalClaimPath = -Join($localPath,'\Json\AadOptionalClaims_Reset.json');
$DEFAULT_SCOPE=$(az ad app show --id $appId | .\jq '.oauth2Permissions[0].isEnabled = false' | .\jq -r '.oauth2Permissions')
$DEFAULT_SCOPE>>scope.json
az ad app update --id $appId --set oauth2Permissions=@scope.json
Remove-Item .\scope.json
az ad app update --id $appId --remove oauth2Permissions
az ad app update --id $appId --set oauth2AllowIdTokenImplicitFlow=false
az ad app update --id $appId --remove replyUrls --remove IdentifierUris
az ad app update --id $appId --identifier-uris "$IdentifierUris"
az ad app update --id $appId --remove IdentifierUris
az ad app update --id $appId --optional-claims $resetOptionalClaimPath
az ad app update --id $appId --remove requiredResourceAccess

# updating reply urls and api permissions
Set-AzureADApplication -ObjectId $appObjectId -ReplyUrls $replyUrls -RequiredResourceAccess $requiredResourceAccess

# setting AAD optional claims:
$localPath = Get-Location
$setOptionalClaimPath = -Join($localPath,'\Json\AadOptionalClaims.json');
az ad app update --id $appId --optional-claims $setOptionalClaimPath
az ad app update --id $appId --set oauth2AllowIdTokenImplicitFlow=true
az ad app update --id $appId --set oauth2AllowImplicitFlow=true

# Step 5 - Taking the object id generated in Step 2, create a new Password
$pwdVars = New-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId

# Step 5a - Updating the logo for the Azure AD app
$location = Get-Location
$imgLocation = -Join($location, '\Assets\moodle-logo.jpg')

Set-AzureADApplicationLogo -ObjectId $app.ObjectId -FilePath $imgLocation

# Step 5b - Add expose an API 
# Reference https://docs.microsoft.com/en-us/answers/questions/29893/azure-ad-teams-dev-how-to-automate-the-app-registr.html
# Expose an API

$msApplication = Get-AzureADMSApplication -ObjectId $appObjectId

# Set identifier URL
$identifierUris = 'api://' + $moodleDomain + '/' + $appId
Set-AzureADMSApplication -ObjectId $msApplication.Id -IdentifierUris $identifierUris
                          
# Create access_as_user scope
$scopes = New-Object System.Collections.Generic.List[Microsoft.Open.MsGraph.Model.PermissionScope]
$msApplication.Api.Oauth2PermissionScopes | foreach-object { $scopes.Add($_) }
$scope = Create-Scope -value "access_as_user"  `
    -userConsentDisplayName "Teams can access the user profile and make requests on the user's behalf"  `
    -userConsentDescription "Enable Teams to call this app’s APIs with the same rights as the user"  `
    -adminConsentDisplayName "Teams can access the user’s profile"  `
    -adminConsentDescription "Allows Teams to call the app’s web APIs as the current user"

$scopes.Add($scope)
$msApplication.Api.Oauth2PermissionScopes = $scopes

Set-AzureADMSApplication -ObjectId $msApplication.Id -Api $msApplication.Api
Write-Host -message "Scope access_as_user added."
             
# Authorize Teams mobile/desktop client and Teams web client to access API
$preAuthorizedApplications = New-Object 'System.Collections.Generic.List[Microsoft.Open.MSGraph.Model.PreAuthorizedApplication]'
$teamsRichClientPreauthorization = Create-PreAuthorizedApplication `
    -applicationIdToPreAuthorize '1fec8e78-bce4-4aaf-ab1b-5451cc387264' `
    -scopeId $scope.Id
$teamsWebClientPreauthorization = Create-PreAuthorizedApplication `
    -applicationIdToPreAuthorize '5e3ce6c0-2b1f-4285-8d4b-75ee78787346' `
    -scopeId $scope.Id
$preAuthorizedApplications.Add($teamsRichClientPreauthorization)
$preAuthorizedApplications.Add($teamsWebClientPreauthorization)

$msApplication.Api.PreAuthorizedApplications = $preAuthorizedApplications
Set-AzureADMSApplication -ObjectId $msApplication.Id -Api $msApplication.Api
Write-Host -message "Teams mobile/desktop and web clients applications pre-authorized."


# Step 6 - Write out the newly generated app Id and azure app password
Write-Host 'Your AD Application ID: '$appId
Write-Host 'Your AD Application Secret: '$pwdVars.Value
