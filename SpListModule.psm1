function Initialize-ModuleData {
    $dataFile = (Import-PowerShellDataFile -Path $PSScriptRoot'\SpListModule.psd1').PrivateData
    #New-Variable -Scope Script -Name 'privateData' -Value $dataFile #-Option ReadOnly
    return $dataFile
}

# connectSharePoint builds a connection to a SharePoint site and imports the module functions
function Connect-SharePoint {
        [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter()][switch]$UseDefaultCredentials
    )
    
    $Script:privateData = Initialize-ModuleData

    # Conditionally set the credential method
    $credentialMethod = if ($UseDefaultCredentials) {
        $UseDefaultCredentials
    }
    else {
        $webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $webSession.Credentials = Get-Credential
        $webSession
    }

    # Store the SharePoint API endpoint list
    $api = $privateData.Api
    $uri = $SiteUrl + $api.contextinfo

    # Build the connection parameters
    $connectionParams = @{
        Uri                   = $uri
        Method                = 'POST'
        SkipHttpErrorCheck    = $true
        UseBasicParsing       = $true
        ContentType           = "application/json;odata=verbose"
    }

    # Conditionally add authentication parameters
    if ($UseDefaultCredentials) { 
        $connectionParams.UseDefaultCredentials = $true 
    }
    else {
        $connectionParams.WebSession = $webSession
    }
    
    try {
        # Get the request digest
        $connection = (Invoke-RestMethod @connectionParams).GetContextWebInformation.FormDigestValue
        
        # Test the connection
        if ($connection) {
            Write-Host "`r`nConnection to $SiteUrl successful" -ForegroundColor Green
        }
        else {
            # This section should probably be a throw
            Write-Host "`r`nConnection to $SiteUrl failed" -ForegroundColor Red
            exit
        }

        # Store the session and other necessary info
        $sessionObj = [PSCustomObject]@{
            SiteUrl        = $SiteUrl
            ListTitle      = $ListTitle
            WebSession     = $credentialMethod
            RequestDigest  = $connection
        }
    }
    catch {
        Write-Host "`r`nConnection to $SiteUrl failed" -ForegroundColor Red
        exit
    }

    # Since we passed the connection test, we can import the module functions
    $importParams = @{
        SessionObj            = $sessionObj
        DefaultProperties     = $privateData.DefaultProperties
        Api                   = $privateData.Api
        UseDefaultCredentials = $UseDefaultCredentials
    }

    # Import the module functions
    . $PSScriptRoot'\SpListModuleFunctions.ps1' @importParams
}

# disconnectSharePoint removes the imported module functions. It could do a lot more.
function Disconnect-SharePoint {
    [CmdletBinding()]
    param (
        [Parameter()][switch]$ShowFunctions
    )

    $importedFunctions = $Script:privateData.ImportedSpListModuleFunctions

    foreach ($function in $importedFunctions) {
        $functionPath = "function:$function"
        if (Test-Path $functionPath) {
            if ($ShowFunctions) {
                Write-Host "Removing $function"
            }

            Remove-Item -Path $functionPath -Force
        }
        else {
            Write-Host "$function does not exist in global scope"
        }
    }
    Write-Host "`r`nDisconnected from $($Script:sessionObj.siteurl)`r`n" -ForegroundColor DarkYellow
}