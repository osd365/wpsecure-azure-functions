if ($env:AZURE_FUNCTIONS_ENVIRONMENT -eq "Development") {
    $env:PSModulePath = "$PSScriptRoot\Modules;$env:PSModulePath"
} else {
    $env:PSModulePath = "/home/site/wwwroot/Modules:$env:PSModulePath"
}