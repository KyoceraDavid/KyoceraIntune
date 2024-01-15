# Define printer name
$printerName = "Kyocera Secure Print"

# Define driver information
$driverName = "Kyocera TASKalfa 4053ci KX"

# Define port information
$portName = "FollowMe Port"

# Check if printer exists and remove if necessary
$existingPrinter = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
if ($existingPrinter) {
    Write-Output "Removing ""$($printerName)""."
    Remove-Printer -Name $printerName
}
else {
    Write-Output "Printer already removed."
}

# Check if printer port exists and remove if necessary
$existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
if ($existingPort) {
    Write-Output "Removing Port ""$($PortName)""."
    Remove-PrinterPort -Name $portName
}
else {
    Write-Output "Port already removed."
}

# Check if printer driver exists and remove if necessary
$existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
if ($existingDriver){
    Write-Output "Removing Print Driver ""$($driverName)""."
    Remove-PrinterDriver -Name $driverName -RemoveFromDriverStore
}
else {
    Write-Output "Print Driver already removed."
}