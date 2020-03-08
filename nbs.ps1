# Adjust to your needs
# ------------------------------------------------------------------------------
$scannerID = 0 # 0-based index of the device, call Print-Scanners if unsure
$dpi = 450
$format = 'tiff' # jpeg, png, tiff
$quality = 95 # jpeg quality
$filePathFormat = 'D:\Home\Documents\From-Scanner\scan-{0:d3}.{1}'
# Custom page extent (x, y, w, h) in mm; eg. @(31, 8, 148, 178); set to $null to disable
$extent = @(31, 8, 148, 178)
# ------------------------------------------------------------------------------

# Exit on any error
$ErrorActionPreference = "Stop"
# WiaDeviceType.ScannerDeviceType
$scannerDeviceType = 1
# FormatIDs
$formats = @{
    'tiff' = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
    'png'  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    'jpeg' = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
}

$formatID = $formats['jpeg']
if ($formats.ContainsKey($format)) {
    $formatID = $formats[$format]
}

function Get-Scanner([Int] $id = $null) {
    $manager = New-Object -ComObject WIA.DeviceManager
    $result = @()
    for ($i = 0; $i -lt $manager.DeviceInfos.Count; $i++) {
        $device = $manager.DeviceInfos.Item($i+1)
        if ($device.Type -eq $scannerDeviceType) {
            $p = $device.Properties
            $result += [PSCustomObject] @{
                ID = $i
                DeviceObject = $device
                DeviceID = $device.DeviceID
                Port = $p.Item('Port').Value
                Name = $p.Item('Name').Value
                Manufacturer = $p.Item('Manufacturer').Value
            }
        }
    }
    if ($id -ge 0) {
        return $result.Item($id)
    }
    return $result
}

function Print-Scanners() {
    foreach ($device in Get-Scanner) {
        Write-Host ($device | Format-Table -Property ID, Manufacturer, Name, Port | Out-String)
    }
}

function Convert-Image($image) {
    $imageProcess = New-Object -ComObject WIA.ImageProcess
    $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
    $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $formatID
    $imageProcess.Filters.Item(1).Properties.Item("Quality").Value = $quality
    return $imageProcess.Apply($image)
}

function Metrics-To-Pixels([Int] $value) {
    return [Math]::Round(($value / 25.4) * $dpi)
}

function Scan() {
    $device = (Get-Scanner $scannerID).DeviceObject.Connect()
    foreach ($item in $device.Items) {
        $item.Properties.Item("Horizontal Resolution").Value = $dpi
        $item.Properties.Item("Vertical Resolution").Value = $dpi
        if ($extent -is [Array]) {
            $item.Properties.Item("Horizontal Start Position").Value = Metrics-To-Pixels($extent[0])
            $item.Properties.Item("Vertical Start Position").Value = Metrics-To-Pixels($extent[1])
            $item.Properties.Item("Horizontal Extent").Value = Metrics-To-Pixels($extent[2])
            $item.Properties.Item("Vertical Extent").Value = Metrics-To-Pixels($extent[3])
        }
        $image = $item.Transfer($formatID)
    }
    if ($image.FormatID -ne $formatID) {
        $image = Convert-Image $image
    }
    return $image
}

function Get-NextPath() {
    $i = 1
    while (Test-Path ($filePathFormat -f $i, $format)) {
        ++$i
    }
    return $filePathFormat -f $i, $format
}

$path = Get-NextPath
New-Item -Path (Split-Path $path -Parent) -ItemType Directory -Force >$null
Write-Host "Scanning to $path ..." -ForegroundColor Yellow
$image = Scan
$image.SaveFile($path)
Write-Host "Done!!" -ForegroundColor Green
