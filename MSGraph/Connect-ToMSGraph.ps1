# Connect to MS Graph via User and Scope
function Connect-ToMSGraph {
    [CmdletBinding(DefaultParameterSetName = 'Connect')]
    param(
        # Parameter set for connecting
        [Parameter(Mandatory=$true, ParameterSetName='Connect')]
        [System.Collections.Generic.List[string]]
        $Scopes,

        [Parameter(Mandatory=$true, ParameterSetName='Connect')]
        [string]
        $TenantId,

        [Parameter(Mandatory=$false, ParameterSetName='Connect')]
        [switch]
        $Silent,

        [Parameter(Mandatory=$false, ParameterSetName='Connect')]
        [System.Management.Automation.PSCredential]
        $Credential,

        # Parameter set for disconnecting
        [Parameter(Mandatory=$true, ParameterSetName='Disconnect')]
        [switch]
        $Disconnect,

        # Common optional switches
        [Parameter(Mandatory=$false, ParameterSetName='Connect')]
        [switch]
        $CheckIfConnected,

        [Parameter(Mandatory=$false, ParameterSetName='Connect')]
        [switch]
        $ForceNewToken
    )
    
    # Load module
    if(-not $(Get-Module Microsoft.Graph.Authentication -ListAvailable)){
        Write-Host "Installing Microsoft Graph PowerShell module..."
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
    }

    if(-not $(Get-Module Microsoft.Graph.Users -ListAvailable)){
        Write-Host "Installing Microsoft Graph Users module..."
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
    }

    # Handle disconnect parameter set
    if ($PSCmdlet.ParameterSetName -eq 'Disconnect') {
        if (-not $Silent) {
            Write-Host "Disconnecting from Microsoft Graph..."
        }
        Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
        if (-not $Silent) {
            Write-Host "Successfully disconnected from Microsoft Graph"
        }
        return $true
    }

    # Handle CheckIfConnected switch early
    if ($CheckIfConnected) {
        try {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($context -and $context.AuthType -eq 'Delegated') {
                Write-Verbose "Graph context exists but is not delegated authentication"
                return $true
            }
            # If we have a context and all our script variables are populated, we're likely connected
            return $false
        }
        catch {
            Write-Verbose "Error checking Graph context: $_"
            return $false
        }

    # If any of the required variables are missing, we're not properly connected
    Write-Verbose "Missing required connection variables"
    return $false

    }
    
    function Get-GraphUserInfo {
        try{        
            # Get token directly from the context to avoid request header extraction issues
            $context = Get-MgContext
            if (-not $context) {
                Write-Error "No Microsoft Graph context found. Please authenticate first."
                return $null
            }
        
            # Extract user info
            try {
                $userInfo = Invoke-GraphRequest -Method GET -Uri "/v1.0/me" -ErrorAction Stop
            }
            catch {
                # Check if error is 404 NotFound (user has no /me resource)
                if ($_.Exception.Response.StatusCode.Value__ -eq 404) {
                    Write-Verbose "User 'me' resource not found, likely unlicensed user. Returning limited info."

                    # Build minimal user info object
                    $userInfo = [PSCustomObject]@{
                        DisplayName       = "Unknown User (No /me resource)"
                        UserPrincipalName = "Unavailable"
                    }
                }
                else {
                    throw $_
                }
            }
        
            # Get access token from context
            $Token = $context.AccessToken
            if (-not $Token) {
                # Fallback to request header extraction if context doesn't have token
                $Parameters = @{
                    Method = "GET"
                    URI = "/v1.0/me"
                    OutputType = "HttpResponseMessage"
                }
                $Response = Invoke-GraphRequest @Parameters
                $Headers = $Response.RequestMessage.Headers
                $Token = $Headers.Authorization.Parameter
            }
        
            Write-Verbose "Successfully retrieved access token for Graph API"
        
            # Add token and expiration info to user object
            $userInfo | Add-Member -NotePropertyName "Token" -NotePropertyValue $Token
            $userInfo | Add-Member -NotePropertyName "ExpiresOn" -NotePropertyValue $context.ExpiresOn
        
            # Return complete user info with token as JSON
            return $userInfo
        }
        catch {
            # Write errors to stderr but return clean empty JSON for the caller
            $_ | Out-String | Write-Error
            return $null
        }
    }

    # Connect all scopes and return token
    try {
        # Check if already connected with required scopes and return early
        $currentContext = Get-MgContext -ErrorAction SilentlyContinue

        # Check expiration
        if ($currentContext -and $currentContext.ExpiresOn -lt (Get-Date).AddMinutes(5)) {
            Write-Verbose "Token expires soon, forcing refresh"
            $null = Invoke-GraphRequest -Method GET -Uri "/v1.0/me" -ErrorAction Stop
            $currentContext = Get-MgContext -ErrorAction SilentlyContinue
        }

        if ($currentContext) {
            $hasAllScopes = $true
            foreach ($scope in $Scopes) {
                if ($currentContext.Scopes -notcontains $scope) {
                    $hasAllScopes = $false
                    break
                }
            }
            
            if ($hasAllScopes) {
                # We're already connected, get user info and token
                try {
                    $userInfo = Get-GraphUserInfo
                    
                    if (-not $Silent) {
                        Write-Host "Connected to Microsoft Graph as $($userInfo.DisplayName) ($($userInfo.UserPrincipalName)) with all scopes requested" -ForegroundColor Green
                    }

                    return $userInfo
                }
                catch {
                    # If we can't get user info, disconnect and try again
                    if (-not $Silent) {
                        Write-Host "Connected but unable to retrieve user information. Reconnecting..." -ForegroundColor Yellow
                    }
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                }
            }
            else {
                # Disconnect if we don't have all required scopes
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }
        }
        # else connect now
        else{
            # Prepare connection parameters
            $connectParams = @{
                Scopes = $Scopes
                TenantId = $TenantId
                ErrorAction = "Stop"
            }
            
            # Add credential if provided
            if ($Credential) {
                $connectParams.Add("Credential", $Credential)
            }
            
            # User info
            if (-not $Silent) {
                Write-Host "Connecting to Microsoft Graph..."
            }
            
            # Try to connect
            try{
                Connect-MgGraph @connectParams | Out-Null
            }
            catch {
                # Properly disconnect
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

                # Clear any cached tokens
                [Microsoft.Identity.Client.PublicClientApplication]::ClearTokenCache()

                # Try to connect again
                Connect-MgGraph @connectParams | Out-Null
            }
                        
            # Get user information
            try {

                $userInfo = Get-GraphUserInfo
                
                if (-not $Silent) {
                    Write-Host "Connected to Microsoft Graph as $($userInfo.DisplayName) $($userInfo.UserPrincipalName)" -ForegroundColor Green
                }

                return $userInfo
            }
            catch {
                # If we can't get user info but are connected, still proceed
                if (-not $Silent) {
                    Write-Host "Connected to Microsoft Graph, but unable to retrieve user details: $_" -ForegroundColor Yellow
                    Write-Host "Connection established with ClientId: $($context.ClientId)" -ForegroundColor Green
                }
            }
            
        }
    }
    catch {
        if (-not $Silent) {
            Write-Host "Authentication failed: $_" -ForegroundColor Red
        }
        return $false
    }
}

#$tokenIinfo = Connect-ToMSGraph -Scopes ("AuditLog.Read.All", "SecurityEvents.Read.All") -TenantId "my-tenant-ID"
#Connect-ToMSGraph -Disconnect
