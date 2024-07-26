param(
    [switch]$uninstall = $false
)

if($uninstall){
    winget uninstall --id Microsoft.Powershell
} else {
    winget install --id Microsoft.Powershell --source winget
}