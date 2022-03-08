# Update Group Policy
function Update-GroupPolicy { 
   
   $gpupdate = "C:\Windows\System32\gpupdate.exe"

   Write-Host "Updating Group Policy..."
   Start-Process -FilePath $gpupdate -Wait -ArgumentList "/force"
   Write-Host "Group Policy successfully updated!" -ForegroundColor Green
   
}

# Add printers
function Map-Printers {

    $printers = @( 
    
    "\\pwgsasmpsjm08\FTWSecure08PCL6.4" 
    "\\pwgsasmpsjm13\TXAM1FL03MM62PS" 
    "\\pwgsasmpsjm13\TXAM2FL01SM07" 
    "\\pwgsasmpsjm13\TXAM2FL01SM08" 
    
    )

    Write-Host "Mapping printers..."
    foreach($printer in $printers) { Add-Printer -ConnectionName $printer; Write-Host "$printer successfully mapped!" -ForegroundColor Green }

}
   
# Map K drive
function Map-Drives {

    $DriveLetter = "K"
    $DrivePath = "\\pwsafshounas13\vdata4\sd1valic"
    
    Write-Host "Mapping Valic K: drive..."
    if (Get-PSDrive $DriveLetter -ErrorAction SilentlyContinue) { 
        Write-Host "The Valic K: drive is already in use." -ForegroundColor Red 
    }
    else { 
        New-PSDrive -Name $DriveLetter -PSProvider "FileSystem" -Root $DrivePath -Persist 
        Write-Host "The Valic K: drive mapped successfully!" -ForegroundColor Green
    }
    
}

# Create desktop shortcuts
function Copy-Shortcuts {
    $shorcuts = "Shortcuts\*"
    $desktop = "$HOME\Desktop"

    Write-Host "Copying shortcuts to desktop..."
    Get-ChildItem -Path $shortcuts | ForEach-Object {Copy-Item $shorcuts -Destination $desktop -Force}
    Write-Host "Shortcuts successfully copied!" -ForegroundColor Green
    
}

Update-GroupPolicy
Map-Printers
Map-Drives
Copy-Shortcuts


