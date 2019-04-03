#### Variables ####

    $computerName = Read-Host -Prompt 'Input computer name'
    $baddevices = Get-WmiObject Win32_PNPEntity | where {$_.ConfigManagerErrorcode -ne 0}
    $colItems = Get-WmiObject -Class Win32_USBHub
    $colSoftware = Get-WmiObject -Class Win32_Product | Where-Object {$_.IdentifyingNumber -eq "{90280409-6000-11D3-8CFE-0050048383C9}"}

### Get Information and Display It ###

    Get-WmiObject -Class Win32_BIOS -ComputerName $computerName
    Get-WmiObject -Class Win32_Processor -ComputerName $computerName
    Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computerName
    Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computerName
    Get-WmiObject Win32_LocalTime -computer $computerName

### Display any malfunctioning devices ###

    " Total Bad devices: {0}" -f $baddevices.count
        foreach ($device in $baddevices) 
            {
            "                               "
            "*******************************"
            "Name : {0}" -f $device.name
            "Class Guid : {0}" -f $device.Classguid
            "Description : {0}" -f $device.Description
            "Device ID : {0}" -f $device.deviceid
            "Manufacturer : {0}" -f $device.manufactuer
            "PNP Devcice Id : {0}" -f $device.PNPDeviceID
            "Service Name : {0}" -f $device.service
            }

### Determines what kind of device is plugged into the USB drives ###

    foreach ($objItem in $colItems)
        {
        "                               "
        "*******************************"
        "Device ID: " + $objItem.DeviceID
        "PNP Device ID: " + $objItem.PNPDeviceID
        "Description: " + $objItem.Description
        }

### Pulls the version of Microsoft Office Installed on the machine ###

    foreach ($colItem in $colSoftware)
        {
        "Name: " + $colItem.Name
        "Version: " + $colItem.Version
        }

### Determines startup programs for machine ###

    $colItems2 = Get-WmiObject -Class Win32_StartupCommand -ComputerName $computerName
    foreach ($objStartupCommand in $colItems2) 
    { 
        "                               "
        "*******************************"
        "Command: " + $objStartupCommand.Command
        "Description: " + $objStartupCommand.Description 
        "Location: " + $objStartupCommand.Location 
        "Name: " + $objStartupCommand.Name 
        "SettingID: " + $objStartupCommand.SettingID 
        "User: " + $objStartupCommand.User
    }