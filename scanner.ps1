while ($true) {

    Clear-Host

    Write-Host "===== IT INVENTORY SCANNER ====="
    Write-Host "1. Show Asset Name"
    Write-Host "2. Show MAC Address"
    Write-Host "3. Show Serial Number"
    Write-Host "4. Show Model"
    Write-Host "5. Scan Device (Auto Add)"
    Write-Host "6. Manually Add Asset"
    Write-Host "7. Exit"
    Write-Host "8. Scan for External Monitors"
    Write-Host ""

    $choice = Read-Host "Please Select an Option"

    switch ($choice) {

        # =========================
        # SHOW ASSET NAME
        # =========================

        "1" {

            Write-Host ""
            Write-Host "Asset Name:"
            Write-Host $env:COMPUTERNAME

            Pause
        }

        # =========================
        # SHOW MAC ADDRESS
        # =========================

        "2" {

            $mac = Get-NetAdapter |
                Where-Object { $_.Status -eq "Up" } |
                Select-Object -First 1 -ExpandProperty MacAddress

            Write-Host ""
            Write-Host "MAC Address:"
            Write-Host $mac

            Pause
        }

        # =========================
        # SHOW SERIAL NUMBER
        # =========================

        "3" {

            $serial = Get-CimInstance Win32_BIOS |
                Select-Object -ExpandProperty SerialNumber

            Write-Host ""
            Write-Host "Serial Number:"
            Write-Host $serial

            Pause
        }

        # =========================
        # SHOW MODEL
        # =========================

        "4" {

            $model = Get-CimInstance Win32_ComputerSystem |
                Select-Object -ExpandProperty Model

            Write-Host ""
            Write-Host "Model:"
            Write-Host $model

            Pause
        }

        # =========================
        # AUTO DEVICE SCAN
        # =========================

        "5" {

            # Asset ID

            do {
                $assetID = Read-Host "Enter Asset ID (or B to go back)"
                if ($assetID -eq "B") { continue 2 }
                if ([string]::IsNullOrWhiteSpace($assetID)) {

                    Write-Host ""
                    Write-Host "Asset ID cannot be empty"
                    Write-Host ""

                }
            }
            while ([string]::IsNullOrWhiteSpace($assetID))

            # Asset Name
            $assetname = $env:COMPUTERNAME

            # User input validation
            do {

                $user = Read-Host "Please enter your name (or B to go back)"

                if ($user -eq "B") { continue 2 }

                if ([string]::IsNullOrWhiteSpace($user)) {

                    Write-Host ""
                    Write-Host "User name cannot be empty"
                    Write-Host ""

                }

            }
            while ([string]::IsNullOrWhiteSpace($user))

            # Location input validation
            do {

                $location = Read-Host "Enter location and patch panel (or B to go back)"

                if ($location -eq "B") { continue 2 }

                if ([string]::IsNullOrWhiteSpace($location)) {

                    Write-Host ""
                    Write-Host "Location cannot be empty"
                    Write-Host ""

                }

            }
            while ([string]::IsNullOrWhiteSpace($location))

            # Date
            $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

            # Model
            $model = Get-CimInstance Win32_ComputerSystem |
                Select-Object -ExpandProperty Model

            # Device Type
            $pcSystemType = Get-CimInstance Win32_ComputerSystem |
                Select-Object -ExpandProperty PCSystemType

            switch ($pcSystemType) {

                1 { $assetType = "Desktop" }
                2 { $assetType = "Laptop" }

                default { $assetType = "Unknown" }
            }

            # Serial
            $serial = Get-CimInstance Win32_BIOS |
                Select-Object -ExpandProperty SerialNumber

            # MAC
            $mac = Get-NetAdapter |
                Where-Object { $_.Status -eq "Up" } |
                Select-Object -First 1 -ExpandProperty MacAddress

            # Comments
            $comments = Read-Host "Enter any additional comments (leave blank if none, or B to go back)"

                if ($comments -eq "B") { continue 2 }

            # Reports Folder
            $reportsFolder = "$PSScriptRoot\reports"

            if (!(Test-Path $reportsFolder)) {

                New-Item -ItemType Directory -Path $reportsFolder | Out-Null
            }

            # CSV
            $csvFile = "$reportsFolder\scanned_assets.csv"

            # DUPLICATE CHECK
            if (Test-Path $csvFile) {

                $existing = Import-Csv $csvFile

                $duplicate = $existing | Where-Object {

                    $_."Asset ID" -eq $assetID
                    $_.SerialNumber -eq $serial
                }

                if ($duplicate) {

                    Write-Host ""
                    Write-Host "WARNING: Device already exists!"

                    Pause
                    continue
                }
            }

            # Create Object
            $device = [PSCustomObject]@{

                "Asset ID" = $assetID
                AssetName = $assetname
                User = $user
                Location = $location
                Date = $date
                AssetType = $assetType
                Model = $model
                SerialNumber = $serial
                MACAddress = $mac
                Comments = $comments
            }

            # Save CSV
            if (Test-Path $csvFile) {

                $device | Export-Csv $csvFile -NoTypeInformation -Append
            }
            else {

                $device | Export-Csv $csvFile -NoTypeInformation
            }

            Write-Host ""
            Write-Host "Device scan saved!"
            Write-Host $csvFile

            Pause
        }

        # =========================
        # MANUAL ASSET ENTRY
        # =========================

        "6" {

            while ($true) {

                Clear-Host

                Write-Host "===== MANUAL ASSET ENTRY ====="
                Write-Host "1. Monitor"
                Write-Host "2. Phone"
                Write-Host "3. Mouse"
                Write-Host "4. Keyboard"
                Write-Host "5. Other"
                Write-Host "B. Back"
                Write-Host ""

                $assetChoice = Read-Host "Select asset type"

                switch ($assetChoice) {

                    "1" { $assettype = "Monitor" }
                    "2" { $assettype = "Phone" }
                    "3" { $assettype = "Mouse" }
                    "4" { $assettype = "Keyboard" }

                    "5" {

                        do {

                            $assettype = Read-Host "Enter asset type (or B to go back)"

                            if ($assettype -eq "B") { continue 3 }

                            if ([string]::IsNullOrWhiteSpace($assettype)) {

                                Write-Host ""
                                Write-Host "Asset type cannot be empty"
                                Write-Host ""

                            }

                        }
                        while ([string]::IsNullOrWhiteSpace($assettype))
                    }

                    "B" {

                        break
                    }

                    default {

                        Write-Host ""
                        Write-Host "Invalid option"

                        Pause
                        continue
                    }
                }

                if ($assetChoice -eq "B") {

                    break
                }

                #asset ID

                do {

                    $assetID = Read-Host "Enter Asset ID (or B to go back)"

                    if ($assetID -eq "B") { continue 2 }

                    if ([string]::IsNullOrWhiteSpace($assetID)) {

                        Write-Host ""
                        Write-Host "Asset ID cannot be empty"
                        Write-Host ""

                    }

                }
                while ([string]::IsNullOrWhiteSpace($assetID))

                # Asset Name
                do {

                    $assetname = Read-Host "Enter asset name (or B to go back)"

                    if ($assetname -eq "B") { continue 2 }

                    if ([string]::IsNullOrWhiteSpace($assetname)) {

                        Write-Host ""
                        Write-Host "Asset name cannot be empty"
                        Write-Host ""

                    }

                }
                while ([string]::IsNullOrWhiteSpace($assetname))

                # Model
                do {

                    $model = Read-Host "Enter model (or B to go back)"

                    if ($model -eq "B") { continue 2 }

                    if ([string]::IsNullOrWhiteSpace($model)) {

                        Write-Host ""
                        Write-Host "Model cannot be empty"
                        Write-Host ""

                    }

                }
                while ([string]::IsNullOrWhiteSpace($model))

                # User
                do {

                    $user = Read-Host "Enter user's name (or B to go back)"

                    if ($user -eq "B") { continue 2 }

                    if ([string]::IsNullOrWhiteSpace($user)) {

                        Write-Host ""
                        Write-Host "User name cannot be empty"
                        Write-Host ""

                    }

                }
                while ([string]::IsNullOrWhiteSpace($user))

                # Location
                do {

                    $location = Read-Host "Enter asset location (or B to go back)"

                    if ($location -eq "B") { continue 2 }

                    if ([string]::IsNullOrWhiteSpace($location)) {

                        Write-Host ""
                        Write-Host "Location cannot be empty"
                        Write-Host ""

                    }

                }
                while ([string]::IsNullOrWhiteSpace($location))

                # Serial (optional)

                $serial = Read-Host "Enter serial number (leave blank if none, or B to go back)"

                if ($serial -eq "B") { continue 2 }

                # Clean blank serials
                if ([string]::IsNullOrWhiteSpace($serial)) {

                    $serial = "N/A"
                }


                # Date
                $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

                # IMEI
                if ($assettype -eq "Phone") {

                    do {

                        $imei = Read-Host "Enter IMEI (or B to go back)"

                        if ($imei -eq "B") { continue 2 }

                        if ([string]::IsNullOrWhiteSpace($imei)) {

                            Write-Host ""
                            Write-Host "IMEI cannot be empty"
                            Write-Host ""

                        }

                    }
                    while ([string]::IsNullOrWhiteSpace($imei))
                }
                else {

                    $imei = ""
                }

                #comments 

                $comments = Read-Host "Enter any additional comments (leave blank if none, or B to go back)"

                if ($comments -eq "B") { continue 2 }

                # Reports Folder
                $reportsFolder = "$PSScriptRoot\reports"

                if (!(Test-Path $reportsFolder)) {

                    New-Item -ItemType Directory -Path $reportsFolder | Out-Null
                }

                # CSV
                $csvFile = "$reportsFolder\manual_assets.csv"

                # DUPLICATE CHECK
                if (Test-Path $csvFile) {

                    $existing = Import-Csv $csvFile

                    $duplicate = $existing | Where-Object {

                        $_."Asset ID" -eq $assetID
                    }

                    if ($duplicate) {

                        Write-Host ""
                        Write-Host "WARNING: Asset with this ID already exists!"
                        Write-Host "Asset ID: $assetID"

                        Pause
                        continue
                    }
                }


                # Create Asset Object
                $asset = [PSCustomObject]@{
                    "Asset ID" = $assetID
                    AssetType = $assettype
                    AssetName = $assetname
                    User = $user
                    Location = $location
                    SerialNumber = $serial
                    IMEI = $imei
                    Model = $model
                    Date = $date
                    Comments = $comments
                }

                # Save
                if (Test-Path $csvFile) {

                    $asset | Export-Csv $csvFile -NoTypeInformation -Append
                }
                else {

                    $asset | Export-Csv $csvFile -NoTypeInformation
                }

                Write-Host ""
                Write-Host "Asset saved successfully!"
                Write-Host $csvFile

                Pause
            }
        }

        # =========================
        # EXIT
        # =========================

        "7" {

            break
        }

        "8" {

            Clear-Host

            Write-Host "===== CONNECTED MONITORS ====="
            Write-Host "Remember to cross referance Scan to physical assets"
            Write-Host ""

            $monitors = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID

            $found = $false

            foreach ($monitor in $monitors) {

                # Convert arrays to readable text
                $manufacturer = -join ($monitor.ManufacturerName | ForEach-Object { [char]$_ })
                $model        = -join ($monitor.UserFriendlyName | ForEach-Object { [char]$_ })
                $serial       = -join ($monitor.SerialNumberID | ForEach-Object { [char]$_ })

                # Clean serial
                if (
                    [string]::IsNullOrWhiteSpace($serial) -or
                    $serial -eq "0" -or
                    $serial -match "^[0\s]*$"
                ) {

                    $serial = "UNKNOWN"
                }

                # Skip likely laptop/internal displays
                if (
                    ![string]::IsNullOrWhiteSpace($model) -and
                    $model -notmatch "Integrated|Internal|Laptop|Generic|Built"
                ) {

                    $found = $true

                    Write-Host "Manufacturer: $manufacturer"
                    Write-Host "Model:        $model"
                    Write-Host "Serial:       $serial"
                    Write-Host ""
                }
            }

            if (-not $found) {

                Write-Host "No external monitors detected."
                Write-Host ""
            }

            Pause
        }

        # =========================
        # INVALID
        # =========================

        default {

            Write-Host ""
            Write-Host "Invalid option"

            Pause
        }
    }
}


