using namespace System.Net
using namespace System.Web

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "A new beacon receive process has started."

#Variable definition

try {
    $upn = $Request.Body.upn
    $hostname = $Request.Body.hostname
    $operatingsystem = $Request.Body.operatingsystem
    $devicemodel = $Request.Body.devicemodel
    $latitude = $Request.Body.latitude
    $longitude = $Request.Body.longitude
    $beaconlaunchtime = $Request.Body.beaconlaunchtime
    $lastusetime = $Request.Body.lastusetime
    $systemboottime = $Request.Body.systemboottime
    $lastlogontime = $Request.Body.lastlogontime
    $bgversion = $Request.Body.bgversion
    $lsversion = $Request.Body.lsversion
    $signewversion = $Request.Body.signewversion
    $sigreplyversion = $Request.Body.sigreplyversion
    $sigwebversion = $Request.Body.sigwebversion
    $scrnsaveversion = $Request.Body.scrnsaveversion
    $teamsbgversion = $Request.Body.teamsbgversion
    $pendingreboot = $Request.Body.pendingreboot
    $compliant = $Request.Body.compliant

    Write-Host "######################################################################"
    Write-Host "A beacon with the following information was received."
    Write-Host "User Principal Name:$upn"
    Write-Host "Device Hostname:$hostname"
    Write-Host "Operating System and Version:$operatingsystem"
    Write-Host "Device Model and Manufacturer:$devicemodel"
    Write-Host "Device Location Latitude:$latitude"
    Write-Host "Device Location Longitude:$longitude"
    Write-Host "Beacon Launch Time:$beaconlaunchtime"
    Write-Host "Device Last Used By User:$lastusetime"
    Write-Host "System Boot Time:$systemboottime"
    Write-Host "User Last Logon Time:$lastlogontime"
    Write-Host "WPSecure Desktop Wallpaper Package Version:$bgversion"
    Write-Host "WPSecure Lock Screen Package Version:$lsversion"
    Write-Host "WPSecure Signature For New Messages Package Version:$signewversion"
    Write-Host "WPSecure Signature For Reply Messages Package Version:$sigreplyversion"
    Write-Host "WPSecure Signature For Web Messages Package Version:$sigwebversion"
    Write-Host "WPSecure Screensaver Package Version:$scrnsaveversion"
    Write-Host "WPSecure Teams Backdrop Package Version:$teamsbgversion"
    Write-Host "Is Device Pending Reboot?:$pendingreboot"
    Write-Host "Is Device Compliant?:$compliant"
    Write-Host "######################################################################"

    #SharePoint ID's stored in Environmental Variables

    $SHAREPOINT_SITE_URL = $env:SHAREPOINT_SITE_URL
    $SHAREPOINT_LIST_NAME = $env:SHAREPOINT_LIST_NAME

    if ((-not [string]::IsNullOrWhiteSpace($SHAREPOINT_SITE_URL)) -and (-not [string]::IsNullOrWhiteSpace($SHAREPOINT_LIST_NAME))) {
        $internalbeaconlaunchtime = try { ([datetime]$beaconlaunchtime).ToUniversalTime().ToString("o") } catch { "1900-01-01T00:00:00Z" }
        $internallastusetime      = try { ([datetime]$lastusetime).ToUniversalTime().ToString("o") } catch { "1900-01-01T00:00:00Z" }
        $internalsystemboottime   = try { ([datetime]$systemboottime).ToUniversalTime().ToString("o") } catch { "1900-01-01T00:00:00Z" }
        $internallastlogontime    = try { ([datetime]$lastlogontime).ToUniversalTime().ToString("o") } catch { "1900-01-01T00:00:00Z" }
        $internallatitude = try { [double]::Parse($latitude) } catch { 0.0 }
        $internallongitude = try { [double]::Parse($longitude) } catch { 0.0 }
        $internalupn = $upn
        $internalhostname = $hostname
        $internalTitle = "$internalupn ON $hostname"
        $internaloperatingsystem = $operatingsystem
        $internaldevicemodel = $devicemodel
        $internalbgversion = $bgversion
        $internallsversion = $lsversion
        $internalsignewversion = $signewversion
        $internalsigreplyversion = $sigreplyversion
        $internalsigwebversion = $sigwebversion
        $internalscrnsaveversion = $scrnsaveversion
        $internalteamsbgversion = $teamsbgversion
        $internalpendingreboot = try { [bool]::Parse($pendingreboot) } catch { $false }
        $internalcompliant = try { [bool]::Parse($compliant) } catch { $false }

        $fieldValues = @{
            "Title" = "$internalTitle"
            "int_upn" = "$internalupn"
            "int_hostname" = "$internalhostname"
            "int_operatingsystem" = "$internaloperatingsystem"
            "int_devicemodel" = "$internaldevicemodel"
            "int_beaconlaunchtime" = "$internalbeaconlaunchtime"
            "int_lastusetime" = "$internallastusetime"
            "int_systemboottime" = "$internalsystemboottime"
            "int_lastlogontime" = "$internallastlogontime"
            "int_bgversion" = "$internalbgversion"
            "int_lsversion" = "$internallsversion"
            "int_signewversion" = "$internalsignewversion"
            "int_sigreplyversion" = "$internalsigreplyversion"
            "int_sigwebversion" = "$internalsigwebversion"
            "int_scrnsaveversion" = "$internalscrnsaveversion"
            "int_teamsbgversion" = "$internalteamsbgversion"
            "int_pendingreboot" = "$internalpendingreboot"
            "int_compliant" = "$internalcompliant"
            "int_location" = "POINT($internallongitude $internallatitude)"
        }

        Connect-PnPOnline -ManagedIdentity -Url $SHAREPOINT_SITE_URL -Debug

        Write-Host "Checking if a record exist for the UPN and HOSTNAME combination."

        $existingItemSearchQuery = "<View><Query><Where><And><Eq><FieldRef Name='int_hostname' /><Value Type='Text'>$internalhostname</Value></Eq><Eq><FieldRef Name='int_upn' /><Value Type='Text'>$internalupn</Value></Eq></And></Where></Query></View>"

        $existingItem = Get-PnPListItem -List $SHAREPOINT_LIST_NAME -Query $existingItemSearchQuery -Debug

        if ($existingItem) {
            Write-Host "Record exist for the UPN and HOSTNAME combination. Existing record will be updated."
            Set-PnPListItem -List $SHAREPOINT_LIST_NAME -Identity $existingItem.Id -Values $fieldValues -Debug
        }
        else {
            Write-Host "Record does not exist for the UPN and HOSTNAME combination. New record will be created."
            Add-PnPListItem -List $SHAREPOINT_LIST_NAME -Values $fieldValues -Debug
        }
    } else {
        Write-Host "The environment variables SHAREPOINT_SITE_URL and/or SHAREPOINT_LIST_NAME are not set."
        # NEW: Return a 400 Bad Request when required configuration is missing
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{ error = "Missing configuration: SHAREPOINT_SITE_URL and/or SHAREPOINT_LIST_NAME are not set." }
            Headers    = @{ "Content-Type" = "application/json" }
        })
        return

    }
}
catch {
    $ex = $_.Exception
    Write-Error "Exception: The process failed. $($ex.Message)"
    Write-Error "Stack Trace: $($ex.StackTrace)" # Uncomment for detailed troubleshooting.
    Write-Error "Inner Exception: $($ex.InnerException)" # Uncomment for detailed troubleshooting.
    
    # Return a 500 Internal Server Error to the caller
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Headers    = @{ "Content-Type" = "application/json" }
        Body       = @{
            error        = "An error occurred while processing the beacon."
            message      = $ex.Message
            # Optional: include diagnostics when you need them
            # stackTrace   = $ex.StackTrace
            # innerMessage = $ex.InnerException?.Message
            # requestId    = $TriggerMetadata?.InvocationId
        }
    })
}
finally {
    Get-PSSession | Remove-PSSession
}