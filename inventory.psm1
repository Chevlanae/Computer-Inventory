class RegEntry {
    [string]$path
    [string]$key
    [string]$value

    RegEntry([hashtable]$value){
        $this.path = $value.path
        $this.key = $value.key
        $this.value = $value.value
    }
    RegEntry([PSCustomObject]$value){
        $this.path = $value.path
        $this.key = $value.key
        $this.value = $value.value
    }

    [bool] IsPresent(){
        if(Get-ItemProperty -Path $this.path | Where-Object -Property $this.key -match $this.value){
            return $true
        } else {
            return $false
        }
    }
}

class RegDetector {
    [RegEntry]$Sophos
    [RegEntry]$Sysaid
    [RegEntry]$Teamviewer

    RegDetector(){
        $uninstallPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"

        $this.Sophos = [RegEntry]::new(@{
            path = $uninstallPath;
            key = "DisplayName";
            value = "Sophos Endpoint Agent";
        })
        $this.Sysaid = [RegEntry]::new(@{
            path = $uninstallPath;
            key = "DisplayName";
            value = "Sysaid Agent";
        })
        $this.Teamviewer = [RegEntry]::new(@{
            path = $uninstallPath;
            key = "DisplayName";
            value = "Teamviewer";
        })
    }

}

class ComputerInfo {
    [string]$name
    [string]$location
    [string]$description
    [string]$deviceType
    [string]$manufacturer
    [string]$model
    [PSCustomObject[]]$MACAddresses
    [string]$serial
    [string]$os
    [string]$owner
    [bool]$inUse
    [bool]$sophos
    [bool]$sysaid
    [bool]$teamviewer

    ComputerInfo(){
        $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        $reg = [RegDetector]::new()

        $this.name = $ComputerSystem.name
        $this.deviceType = if(Get-CimInstance -ClassName Win32_Battery){ "Laptop" } else { "Workstation" }
        $this.manufacturer = $ComputerSystem.manufacturer
        $this.model = $ComputerSystem.Model
        $this.MACAddresses = Get-CimInstance win32_networkadapterconfiguration | Select-Object description, macaddress 
        $this.serial = (Get-CimInstance -class win32_bios).SerialNumber
        $this.os = ((Get-CimInstance -class Win32_OperatingSystem).Name -split "\|")[0]
        $this.sophos = $reg.Sophos.isPresent()
        $this.sysaid = $reg.Sysaid.isPresent()
        $this.teamviewer = $reg.Teamviewer.isPresent()
    }
    ComputerInfo([hashtable]$value){
        $this.name = $value.name
        $this.location = $value.location
        $this.description = $value.description
        $this.deviceType = $value.deviceType
        $this.manufacturer = $value.manufacturer
        $this.MACAddresses = $value.MACAddresses
        $this.serial = $value.serial
        $this.os = $value.os
        $this.owner = $value.owner
        $this.inUse = $value.inUse
        $this.sophos = $value.sophos
        $this.sysaid = $value.sysaid
        $this.teamviewer = $value.teamviewer
    }

    ComputerInfo([PSCustomObject]$value){
        $this.name = $value.name
        $this.location = $value.location
        $this.description = $value.description
        $this.deviceType = $value.deviceType
        $this.manufacturer = $value.manufacturer
        $this.MACAddresses = $value.MACAddresses
        $this.serial = $value.serial
        $this.os = $value.os
        $this.owner = $value.owner
        $this.inUse = $value.inUse
        $this.sophos = $value.sophos
        $this.sysaid = $value.sysaid
        $this.teamviewer = $value.teamviewer
    }

    [string] MACString(){
        $stringArr = $this.MACAddresses | ForEach-Object {"$($_.description): $($_.macaddress)"}
        return  ($stringArr -join ", ")
    }

    [PSCustomObject] GenerateExcelRow(){
        return [PSCustomObject]@{
            "Name" = $this.name
            "Location" = $this.location
            "Description" = $this.description
            "Device Type" = $this.deviceType
            "Manufacturer" = $this.manufacturer
            "MAC Addresses" = $this.MACString()
            "Serial" = $this.serial
            "OS" = $this.os
            "Owner" = $this.owner
            "In Use?" = $this.inUse
            "Sophos?" = $this.sophos
            "Sysaid?" = $this.sysaid
            "Teamviewer?" = $this.teamviewer
            "Correct groups?" = ""
        }
    }
}