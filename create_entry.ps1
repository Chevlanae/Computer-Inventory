using module .\inventory.psm1

if(-not (Get-Command Import-Excel)){
    Install-Module -Name ImportExcel -Force
}

$sites = ("&COMO", "&CAMO", "&COPRO", "&MOBO")

$computerInfo = [ComputerInfo]::new()

$computerInfo.location = Read-Host -Prompt "Enter the general location of the machine (room, department, etc.)"
$computerInfo.description = Read-Host -Prompt "Enter a general description of the machine (purpose, personal or general use, etc.)"
$computerInfo.inUse = $host.UI.PromptForChoice("", "Is this machine currently in use?", ("&No", "&Yes"), 1)

$sheetName = $sites[$host.UI.PromptForChoice('', "Enter the location code to create an entry for", $sites, 0)].Replace("&", "") + "_Inventory.xlsx"
$sheetPath = "$PSScriptRoot\$sheetName"

if(-not [System.IO.File]::Exists($sheetPath)){
    New-Item -Name $sheetName -Path $PSScriptRoot -ItemType File
}

$data = Import-Excel $sheetPath
$duplicateEntry = ($false, 0)

Write-Host $data

$i = 0
foreach($item in $data){
    if($item.Serial -eq $computerInfo.serial){
        Write-Host "Duplicate!!!"
        $duplicateEntry = ($true, $i)
    }
    $i++
}

if($duplicateEntry[0]){
    switch($host.UI.PromptForChoice('', "Entry with serial `"$($computerInfo.serial)`" already exists. Replace entry?", ("&No", "&Yes"), 0)){
        0 {
            Write-Host "No entry created. Exiting script." -ForegroundColor Red
        }
        
        1 {
            if($data.Count -lt 2) {
                $data = $computerInfo.GenerateExcelRow()
            } else {
                $data[$duplicateEntry[1]] = $computerInfo.GenerateExcelRow()
            }
        }
    }
} elseif($data.Count -lt 1) {
    $data = $computerInfo.GenerateExcelRow()
} elseif(($data.GetType().Name) -eq "PSCustomObject"){
    $data = @($data, $computerInfo.GenerateExcelRow())
} else {
    $data += $computerInfo.GenerateExcelRow()
}

$data | Export-Excel $sheetPath
