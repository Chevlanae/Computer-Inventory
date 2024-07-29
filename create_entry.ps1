using module .\inventory.psm1

$sites = ("&COMO", "&CAMO", "&COPRO", "&MOBO")

$computerInfo = [ComputerInfo]::new()

$computerInfo.location = Read-Host -Prompt "Enter the general location of the machine (room, department, etc.)"
$computerInfo.description = Read-Host -Prompt "Enter a general description of the machine (purpose, personal or general use, etc.)"
$computerInfo.inUse = $host.UI.PromptForChoice("", "Is this machine currently in use?", ("&No", "&Yes"), 1)
if($computerInfo.inUse){
    $computerInfo.owner = Read-Host -Prompt "Who is the owner/primary user of this machine?"
}

$sheetName = $sites[$host.UI.PromptForChoice('', "Enter the location code to create an entry for", $sites, 0)].Replace("&", "") + "_Inventory.csv"
$sheetPath = "$PSScriptRoot\$sheetName"

if(-not [System.IO.File]::Exists($sheetPath)){
    New-Item -Name $sheetName -Path $PSScriptRoot -ItemType File
}

$data = Import-Csv -Path $sheetPath
$duplicateEntry = ($false, 0)

if($data -and ($data.GetType()).Name -eq "Object[]"){
    $i = 0
    foreach($item in $data){
        if($item.Serial -eq $computerInfo.serial){
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
                $data[$duplicateEntry[1]] = $computerInfo.GenerateExcelRow()
            }
        }
    } else {
        $data += $computerInfo.GenerateExcelRow()
    }
    
    

} 

elseif($data -and ($data.GetType()).Name -eq "PSCustomObject"){
    $data = @($data, $computerInfo.GenerateExcelRow())   
}

else {
    $data = $computerInfo.GenerateExcelRow()
}




Remove-Item -Path $sheetPath -Force
$data | Select-Object -Property * | Export-Csv -path $sheetPath
