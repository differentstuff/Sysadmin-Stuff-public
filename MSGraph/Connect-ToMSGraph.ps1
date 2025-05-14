function Connect-ToMSGraph {
    [CmdletBinding()]
    param(
        # Define required scopes for OneDrive access
        [Parameter(Mandatory=$true)]
        [System.Collections.Generic.List[string]]
        $Scopes,
        
        [Parameter(Mandatory=$true)]
        [string]
        $TenantId,
        
        [Parameter(Mandatory=$false)]
        [switch]
        $Silent,
        
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter(Mandatory=$false)]
        [switch]$Disconnect,
        
        [Parameter(Mandatory=$false)]
        [switch]$CheckIfConnected,
        
        [Parameter(Mandatory=$false)]
        [switch]$ForceNewToken
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

    # Check if connection exists
    if($CheckIfConnected){
        # Check if we have a valid Graph context
        try {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($context -and $context.AuthType -ne 'Delegated') {
                Write-Verbose "Graph context exists but is not delegated authentication"
                return $false
            }
            
            # If we have a context and all our script variables are populated, we're likely connected
            return $true
        }
        catch {
            Write-Verbose "Error checking Graph context: $_"
            return $false
        }
    
    # If any of the required variables are missing, we're not properly connected
    Write-Verbose "Missing required connection variables"
    return $false
    }

    # Disconnect all
    if($Disconnect){

        if(-not $Silent) {
            Write-Host "Disconnecting from Microsoft Graph..."
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

        if(-not $Silent) {
            Write-Host "Successfully disconnected from Microsoft Graph"
        }

        return $true
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
                    $userInfo = Get-OnedriveInfoAndToken
                    
                    if (-not $Silent) {
                        Write-Host "Connected to Microsoft Graph as $($userInfo.DisplayName) $($userInfo.UserPrincipalName) with required scopes" -ForegroundColor Green
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
                $userInfo = Get-OnedriveInfoAndToken
                
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