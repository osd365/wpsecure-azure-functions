
using namespace System.Net
using namespace System.Web

param($Request, $TriggerMetadata)

Write-Host "A new process has started to get User attributes from EntraID using Microsoft Graph."

try {
    
    # Get UPN from body; fall back to query string
    $upn = $Request.Body.upn
    if ([string]::IsNullOrEmpty($upn)) { $upn = $Request.Query.upn }

    Write-Output "The UPN of the authenticated user is: $upn"
    if (-not [string]::IsNullOrEmpty($upn)) {
        Connect-MgGraph -Identity -NoWelcome

        # Default attributes (used if client doesn't supply requiredAttribs)
        $defaultAttribs = @(
            "businessPhones",
            "city",
            "companyName",
            "country",
            "department",
            "displayName",
            "employeeId",
            "faxNumber",
            "givenName",
            "jobTitle",
            "mail",
            "mailNickname",
            "mobilePhone",
            "officeLocation",
            "postalCode",
            "state",
            "streetAddress",
            "surname",
            "interests",
            "userPrincipalName"
        )

        # Prefer attributes from request body if present and valid
        $incomingAttribs = $Request.Body.requiredAttribs
        if ($incomingAttribs -and $incomingAttribs -is [System.Collections.IEnumerable] -and -not ($incomingAttribs -is [string])) {
            # Normalize to unique strings and trim whitespace
            $requiredAttribs = @()
            foreach ($a in $incomingAttribs) {
                if ($null -ne $a) {
                    $requiredAttribs += $a.ToString().Trim()
                }
            }
            # Distinct, remove empties
            $requiredAttribs = $requiredAttribs | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique
        }
        else {
            $requiredAttribs = $defaultAttribs
        }

        # Guard: Graph requires at least one property
        if (-not $requiredAttribs -or $requiredAttribs.Count -eq 0) {
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Headers    = @{ "Content-type" = "application/json" }
                Body       = @{ message = "No valid requiredAttribs provided." } | ConvertTo-Json
            })
            return
        }

        # Query Microsoft Graph API for user details
        $user = Get-MgUser -UserId $upn -Property ($requiredAttribs -join ',')

        if (-not $user) {
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::NotFound
                Headers    = @{ "Content-type" = "application/json" }
                Body       = @{ message = "No user found for UPN '$upn'." } | ConvertTo-Json
            })
            return
        }

        # Ordered hashtable to preserve output order as per requiredAttribs
        $userDetails = [ordered]@{}

        function Test-IsEnumerableNonString {
            param([object]$Value)
            return ($null -ne $Value -and
                    $Value -is [System.Collections.IEnumerable] -and
                    -not ($Value -is [string]))
        }

        # Auto-populate from requiredAttribs
        foreach ($propName in $requiredAttribs) {
            $prop = $user.PSObject.Properties[$propName]
            if (-not $prop) { continue }

            $value = $prop.Value
            if (Test-IsEnumerableNonString $value) {
                $i = 1
                foreach ($item in $value) {
                    $s = if ($null -ne $item) { $item.ToString() } else { '' }
                    $userDetails["$propName$i"] = $s
                    $i++
                }
            }
            else {
                $userDetails[$propName] = if ($null -ne $value) { $value.ToString() } else { '' }
            }
        }

        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Headers    = @{ "Content-type" = "application/json" }
            Body       = $userDetails | ConvertTo-Json -Depth 5
        })
    }
    else {
        Write-Host "Error: UPN not identified."
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Headers    = @{ "Content-type" = "application/json" }
            Body       = @{ message = "UPN is required in request body or query string." } | ConvertTo-Json
        })
    }
}
catch {
    $ex = $_.Exception
    Write-Host "Exception: The process failed. $($ex.Message)"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Headers    = @{ "Content-type" = "application/json" }
        Body       = @{ message = "Exception: The process failed. $($ex.Message)" } | ConvertTo-Json
    })
}
