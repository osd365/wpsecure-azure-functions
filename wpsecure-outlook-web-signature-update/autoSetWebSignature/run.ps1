using namespace System.Net
using namespace System.Web

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "A new process has started to set Exchange email signature for a User."

# Implement additional validation logic for Exchange Online signature update requests here.
# Verify that the UPN extracted from the client certificate matches the UPN supplied by the calling device.
# Inspect and validate certificate attributes to ensure the request originates from a trusted and compliant source.
# The code below provides a functional baseline; consider extending it with further security controls to harden this Azure Function.

try {
    $EXCHANGE_ORGANIZATIONAL_UNIT_ROOT = $env:EXCHANGE_ORGANIZATIONAL_UNIT_ROOT
    if (-not [string]::IsNullOrEmpty($EXCHANGE_ORGANIZATIONAL_UNIT_ROOT) -and $EXCHANGE_ORGANIZATIONAL_UNIT_ROOT -ne "") {
        # Get signature related environment variables
        $AUTO_ADD_SIGNATURE_ON_NEW_MESSAGE = $true
        $AUTO_ADD_SIGNATURE_ON_REPLY_MESSAGE = $true
        try {$AUTO_ADD_SIGNATURE_ON_NEW_MESSAGE = [bool]::Parse($env:AUTO_ADD_SIGNATURE_ON_NEW_MESSAGE) } catch {$AUTO_ADD_SIGNATURE_ON_NEW_MESSAGE = $true} # Set as detault signature for New messages?
        try {$AUTO_ADD_SIGNATURE_ON_REPLY_MESSAGE = [bool]::Parse($env:AUTO_ADD_SIGNATURE_ON_REPLY_MESSAGE) } catch {$AUTO_ADD_SIGNATURE_ON_REPLY_MESSAGE = $true} # Set as detault signature for Reply messages?
        Import-Module ExchangeOnlineManagement -Force
        Connect-ExchangeOnline -ManagedIdentity -Organization $EXCHANGE_ORGANIZATIONAL_UNIT_ROOT
        $upn = $Request.Body.upn;
        if (-not [string]::IsNullOrEmpty($upn) -and $upn -ne "") {
            $signatureHtml = [System.Web.HttpUtility]::HtmlDecode($Request.Body.signatureHtml)
            $signatureText = [System.Web.HttpUtility]::HtmlDecode($Request.Body.signatureText);
            Write-Host "Starting the process to set Exchange email signature for mailbox with UPN => $upn."
            if ((-not [string]::IsNullOrEmpty($signatureHtml)) -and ($signatureHtml -ne "") -and (-not [string]::IsNullOrEmpty($signatureText)) -and ($signatureText -ne "")) {
                Write-Host "Setting HTML and TXT signatures for UPN => $upn."
                Set-MailboxMessageConfiguration -Identity $upn -SignatureHTML $signatureHtml -SignatureText $signatureText
                Write-Host "Finished setting HTML and TXT signature for UPN => $upn."
            } elseif ((-not [string]::IsNullOrEmpty($signatureHtml)) -and ($signatureHtml -ne "")) {
                Write-Host "Setting HTML signature for UPN => $upn."
                Set-MailboxMessageConfiguration -Identity $upn -SignatureHTML $signatureHtml
                Write-Host "Finished setting HTML signature for UPN => $upn."
            } elseif ((-not [string]::IsNullOrEmpty($signatureText)) -and ($signatureText -ne "")) {
                Write-Host "Setting TXT signature for UPN => $upn."
                Set-MailboxMessageConfiguration -Identity $upn -SignatureText $signatureText
                Write-Host "Finished setting TXT signature for UPN => $upn."
            } else {
                Write-Host "Signature files empty or NULL. Did not set HTML and TXT signatures for UPN => $upn."
            }
            Write-Host "Setting AUTO signature application settings for UPN => $upn."
            Set-MailboxMessageConfiguration -Identity $upn -AutoAddSignature $AUTO_ADD_SIGNATURE_ON_NEW_MESSAGE -AutoAddSignatureOnReply $AUTO_ADD_SIGNATURE_ON_REPLY_MESSAGE
            Write-Host "Finished setting AUTO signature application settings for UPN => $upn."
        }
        else {
            Write-Host "Error: UPN not identified."
            # NEW: Return a 400 Bad Request when required configuration is missing
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{ error = "Error: UPN not identified." }
            Headers    = @{ "Content-Type" = "application/json" }
        })
        return
        }
    } else {
        Write-Host "Error: No EXCHANGE_ORGANIZATIONAL_UNIT_ROOT value supplied for this transaction."
        # NEW: Return a 400 Bad Request when required configuration is missing
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{ error = "Error: No EXCHANGE_ORGANIZATIONAL_UNIT_ROOT value supplied for this transaction." }
            Headers    = @{ "Content-Type" = "application/json" }
        })
        return
    }
}
catch {
    $ex = $_.Exception
    Write-Host "Exception: The process failed. $($ex.Message)"
    Write-Error "Stack Trace: $($ex.StackTrace)" # Uncomment for detailed troubleshooting.
    Write-Error "Inner Exception: $($ex.InnerException)" # Uncomment for detailed troubleshooting.
    
    # Return a 500 Internal Server Error to the caller
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Headers    = @{ "Content-Type" = "application/json" }
        Body       = @{
            error        = "An error occurred while processing the Outlook web signature update request."
            message      = $ex.Message
            # Optional: include diagnostics when you need them
            # stackTrace   = $ex.StackTrace
            # innerMessage = $ex.InnerException?.Message
            # requestId    = $TriggerMetadata?.InvocationId
        }
    })
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false
    Get-PSSession | Remove-PSSession
}