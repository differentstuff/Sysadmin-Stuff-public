function Get-TokenInfo {
    try{        
        # Get token directly from the context to avoid request header extraction issues
        $context = Get-MgContext
        if (-not $context) {
            Write-Error "No Microsoft Graph context found. Please authenticate first."
            return "{}"
        }
        
        # Extract user info
        $userInfo = Get-MgUser -UserId "me" -ErrorAction Stop
        
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
        return $userInfo | ConvertTo-Json -Compress -Depth 10 -ErrorAction Stop
    }
    catch {
        # Write errors to stderr but return clean empty JSON for the caller
        $_ | Out-String | Write-Error
        return "{}"
    }
}
