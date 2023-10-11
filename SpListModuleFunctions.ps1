<#
    This file contains the functions used by the SpListModule.psm1 file.
    In this first release, the functions do not include appropriate help documentation.
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [hashtable]$Api,
    
    [Parameter(Mandatory)]
    [string[]] $DefaultProperties, 
    
    [Parameter(Mandatory)]
    [pscustomobject]$SessionObj,
    
    [switch]$UseDefaultCredentials
)

# Set the received parameters to $Script:variables>
$Script:api = $Api                                      # This is a list of the SharePoint API endpoints
$Script:defaultProperties = $DefaultProperties          # This is a list of some of the default properties that are returned by the SharePoint API
$Script:sessionObj = $SessionObj                        # This is the object that contains the session information
$Script:useDefaultCredentials = $UseDefaultCredentials  # This is a switch that determines whether to use default credentials or not

# Invoke-SharePointApi packages the Invoke-RestMethod and Invoke-WebRequest cmdlets
function Script:Invoke-SharePointApi {
    param(
        [Parameter(Mandatory)]
        [string]$ApiPath,
        
        [Parameter(Mandatory)]
        [ValidateSet('Get','Post','Patch','Delete')]
        [string]$Method,
        
        [switch]$VerboseMode,
        
        [switch]$RequestDigest
    )

    <#
        The $VerboseMode switch is not implemented well throughout the function list.
        Add it to the todo list to better incorporate it for the next release.
        And maybe find a better name for it.
    #>

    # Build the request parameters:
    $requestUrl  = $Script:sessionObj.SiteUrl + $ApiPath
    $requestParams = @{
        Uri                   = $requestUrl
        Method                = $Method
        SkipHttpErrorCheck    = $true
    }

    # Conditionally add authentication parameters
    if ($UseDefaultCredentials) { $requestParams.UseDefaultCredentials = $true }
    else { $requestParams.WebSession = $Script:sessionObj.WebSession }

    # Conditionally add properties based on the HTTP request method
    if ($Method -ne 'Get' -or $RequestDigest) {
        # Build the request headers:
        $headers = @{
            "Accept"          = "application/json;odata=verbose"
            "Content-Type"    = "application/json"
        }

        # Conditionally add the request digest and If-Match headers
        if (-not $RequestDigest) {
            $headers."X-RequestDigest"     = Global:Get-SharePointRequestDigest
            $headers."If-Match"            = "*"
            $requestParams.Headers         = $headers
            $requestParams.Body            = $Script:jsonBody
            $requestParams.UseBasicParsing = $true
        }
    }

    # Send the request:
    try {
        # Conditionally invoke the appropriate cmdlet
        if ($Method -eq "Get" -or $RequestDigest) {
            $request = Invoke-RestMethod @requestParams
            return $request
        }
        # Invoke-RestMethod does too much pre-processing for POST, PATCH, and DELETE requests. Use Invoke-WebRequest instead.
        else {
            $request = Invoke-WebRequest @requestParams

            # VerboseMode is a holdover from the original script. It is not implemented well throughout the function list.
            if ($VerboseMode -and $request.StatusCode -like "2*") {
                Write-Host "`nResponse: " -ForegroundColor Green -NoNewline
                Write-Host "$($request.StatusCode) - " -ForegroundColor Green -NoNewline
                Write-Host "$($request.StatusDescription)" -ForegroundColor Green
            }
            elseif ($request.StatusCode -notlike "2*") {
                Write-Host "`nResponse: " -ForegroundColor Red -NoNewline
                Write-Host "$($request.StatusCode) - " -ForegroundColor Red -NoNewline
                Write-Host "$($request.StatusDescription)" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "`r`nError while invoking SharePoint API:" -ForegroundColor Red -NoNewline
        throw $_.Exception.Message
    }
}

# Convert-SharePointProperties unwraps the XML on a few properties and adds them to the return object
function Script:Convert-SharePointProperties {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Xml.XmlElement] $ListItem
    )

    process
    {
        foreach ($property in $Script:defaultProperties){
            $extroplateParams = @{
                InputObject = $ListItem
                NotePropertyName = "Extended_$property"
                # Forgive me padre, for I have sinned. I love the ternary operator.
                NotePropertyValue = ($ListItem.$property.count -gt 1) ? $ListItem.$property.innertext[0] : $ListItem.$property.innertext
            }

            Add-Member @extroplateParams
        }
    }
}

# Get-SharePointUserId gets the user's ID from the sharepoint site.
function Global:Get-SharePointUserId {
    [CmdletBinding(SupportsShouldProcess = $false)]
    param (
        [Parameter(Mandatory, ValueFromPipeline)][string] $UserEmail
    )
    begin {
        $restParams = @{
            Method  = "Get"
        }
    }
    process {
        $restParams.ApiPath = $Script:api.getByEmail -f $UserEmail
        # Get the user's ID:
        return (Script:Invoke-SharePointApi @restParams).entry.content.properties.id.'#text'
    }
}

# Get-SharePointUser gets the user's DisplayName from the sharepoint site by ID.
function Global:Get-SharePointUser {
    # This function will get the user's DisplayName from the sharepoint site by ID.
    [CmdletBinding(SupportsShouldProcess = $false)]
    param (
        [Parameter(Mandatory, ValueFromPipeline)][int]$UserID
    )
    begin {
        $restParams = @{
            Method  = "Get"
        }
    }
    process {
        $restParams.ApiPath = $Script:api.getUserById -f $UserID
        # Get the user's DisplayName:
        return (Script:Invoke-SharePointApi @restParams).entry.content.properties.Title
    }
}

# Get-SharePointListItems gets the list items from a sharepoint list.
function Global:Get-SharePointListItems {

    # Build the request URL using the $top parameter to get all the items:
    $allItems = @()
    $skipToken = $null
    $top = '?$top=5000'
    $restParams = @{ Method  = "Get" }

    # Get the list items. Originally implemented as a recursive function. Might go back to that. /shrug
    do {
        $restParams.ApiPath = $Script:api.getByTitle -f $Script:sessionObj.ListTitle + $top
        if ($skipToken) {
            $restParams.ApiPath += $Script:api.skipToken -f $skipToken
        }

        $response = (Script:Invoke-SharePointApi @restParams).content.properties
        $allItems += $response

        if ($response.content.properties.__next) {
            $skipToken = $response.content.properties[-1].Id."#text"[0]
        }
        else {
            $skipToken = $null
        }
    } while ($skipToken)

    <#
        For some reason, the last item in the array is sometimes null.
        This will remove the last item if it is null.
    #>
    if ($null -eq $allItems[-1]) {
        $allItems = $allItems[0..($allItems.Length - 2)]
    }

    $allItems | Script:Convert-SharePointProperties
    return $allItems
}

# Get-SharePointAllLists gets all the lists from a sharepoint site.
function Global:Get-SharePointAllLists {
    [CmdletBinding()]
    param (
        [Parameter()][switch]$Raw
    )

    # Use $top to get all the lists (as it's highly unlikely there will be more than 5000):
    $top = '?$top=5000'
    $restParams = @{
        ApiPath = $Script:api.lists + $top
        Method  = "Get"
    }

    # Get the list items:
    $response   = Script:Invoke-SharePointApi @restParams

    $results = $Raw ? $response.content.properties : $response.content.properties.Title

    return $results
}

# Get-SharePointRequestDigest gets the request digest from a sharepoint site.
function Global:Get-SharePointRequestDigest {

    $restParams = @{
        ApiPath       = $Script:api.contextinfo
        Method        = "Post"
        RequestDigest = $true
    }

    return (Script:Invoke-SharePointApi @restParams).GetContextWebInformation.FormDigestValue

}

<#
    The following C_UD functions could use some work.
    VerboseMode does not provide much value beyond telling you whether the request was successful or not.
    It would be nice to have a switch that would allow you to see the request and response bodies.
    Error handling is all done within the Invoke-SharePointApi function. Which is fine, 
    But it would be nice to have some more granular error handling.
#>


# New-SharepointListItem creates a list item in a sharepoint list.
function Global:New-SharepointListItem {
    param(
        [Parameter(Mandatory)]$RequestBody,
        [Parameter()][switch]$VerboseMode
    )

    $Script:jsonBody = $RequestBody | ConvertTo-Json

    $createParams = @{
        ApiPath     = $Script:api.createListItem -f $Script:sessionObj.ListTitle
        Method      = "Post"
        VerboseMode = $VerboseMode
    }

    Script:Invoke-SharePointApi @createParams
}

# Update-SharepointListItem updates a list item in a sharepoint list.
function Global:Update-SharepointListItem {
    param(
        [Parameter(Mandatory)][int]$ItemId,
        [Parameter(Mandatory)]$RequestBody,
        [Parameter()][switch]$VerboseMode
    )

    $Script:jsonBody = $RequestBody | ConvertTo-Json

    $updateParams = @{
        ApiPath     = $Script:api.updateListItem -f $Script:sessionObj.ListTitle, $ItemId
        Method      = "Patch"
        VerboseMode = $VerboseMode
    }

    Script:Invoke-SharePointApi @updateParams
}

# Remove-SharepointListItem deletes a list item in a sharepoint list.
function Global:Remove-SharepointListItem {
    param(
        [Parameter(Mandatory)][int]$ItemId,
        [Parameter()][switch]$VerboseMode
    )

    $deleteParams = @{
        ApiPath     = $Script:api.deleteListItem -f $Script:sessionObj.ListTitle, $ItemId
        Method      = "Delete"
        VerboseMode = $VerboseMode
    }

    Script:Invoke-SharePointApi @deleteParams
}