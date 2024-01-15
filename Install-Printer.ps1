# Define printer name
$printerName = "Kyocera Secure Print"

# Define driver information
$driverName = "Kyocera TASKalfa 4053ci KX"
$driverInfPath = "KyoceraDriver\Kx84_UPD_8.4.1716_en_RC5_WHQL\64bit\OEMSETUP.INF"
$driverStorePath = "C:\Windows\System32\DriverStore\FileRepository\oemsetup.inf_amd64_9359bc628c627b7a\OEMSETUP.INF"

# Define port information
$portName = "FollowMe Port"
$ipAddress = "10.1.1.1"
$lprQueueName = "FollowMe"

# Check if the printer driver exists
$existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
if ($existingDriver) {
    Write-Output "Print Driver ""$($driverName)"" already exists. Removing old driver..."
    Remove-PrinterDriver -Name $driverName -RemoveFromDriverStore -ErrorAction SilentlyContinue
}

# Add Printer Driver
Write-Output "Adding Printer Driver ""$($driverName)"""
C:\Windows\SysNative\pnputil.exe /a $driverInfPath
Add-PrinterDriver -Name $driverName -InfPath $driverStorePath

# Check if the printer port exists
$existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
if ($existingPort) {
    Write-Output "Port ""$($PortName)"" already exists. Removing old port..."

    # Check if the port is associated with a printer and remove it
    $associatedPrinter = Get-WmiObject Win32_Printer | Where-Object { $_.PortName -eq $portName }
    if ($associatedPrinter) {
        Write-Output "Associated printer found. Removing associated printer..."
        Remove-Printer -Name $associatedPrinter.Name -ErrorAction SilentlyContinue
    }

    Remove-PrinterPort -Name $portName -ErrorAction SilentlyContinue
}

# Add Printer Port
Write-Output "Adding Port ""$($portName)"""
Add-PrinterPort -Name $portName -LprHostAddress $ipAddress -LprQueueName $lprQueueName -LprByteCounting

# Check if the printer exists
$existingPrinter = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
if ($existingPrinter) {
    Write-Output "Printer ""$($printerName)"" already exists. Removing old printer..."
    Remove-Printer -Name $printerName -ErrorAction SilentlyContinue
}

# Add Printer
Write-Output "Adding Printer ""$($printerName)"""
Add-Printer -Name $printerName -DriverName $driverName -PortName $portName
