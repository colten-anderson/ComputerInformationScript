#### Variables ####

    $baddevices = Get-WmiObject Win32_PNPEntity | where {$_.ConfigManagerErrorcode -ne 0}
    $colItems = Get-WmiObject -Class Win32_USBHub
    $colSoftware = Get-WmiObject -Class Win32_Product | Where-Object {$_.IdentifyingNumber -eq "{90280409-6000-11D3-8CFE-0050048383C9}"}
	
### Opens input form for computer name ###

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = 'Data Entry Form'
	$form.Size = New-Object System.Drawing.Size(300,200)
	$form.StartPosition = 'CenterScreen'

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Point(75,120)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = 'OK'
	$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $OKButton
	$form.Controls.Add($OKButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Point(150,120)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = 'Cancel'
	$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $CancelButton
	$form.Controls.Add($CancelButton)

	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10,20)
	$label.Size = New-Object System.Drawing.Size(280,20)
	$label.Text = 'Please enter the Computer Name in the space below:'
	$form.Controls.Add($label)

	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Location = New-Object System.Drawing.Point(10,40)
	$textBox.Size = New-Object System.Drawing.Size(260,20)
	$form.Controls.Add($textBox)

	$form.Topmost = $true

	$form.Add_Shown({$textBox.Select()})
	$result = $form.ShowDialog()

	if ($result -eq [System.Windows.Forms.DialogResult]::OK)
	{
		$computerName = $textBox.Text
	}

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