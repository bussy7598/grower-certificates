# Run this script as Administrator!

# 1. Uninstall Python programs via registry
Write-Host "Uninstalling Python programs from registry..."
$uninstallKeys = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
)

foreach ($key in $uninstallKeys) {
    $subkeys = Get-ChildItem -Path $key -ErrorAction SilentlyContinue
    foreach ($subkey in $subkeys) {
        $displayName = (Get-ItemProperty -Path $subkey.PSPath -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
        if ($displayName -and $displayName -match "Python") {
            Write-Host "Uninstalling: $displayName"
            $uninstallString = (Get-ItemProperty -Path $subkey.PSPath -Name UninstallString -ErrorAction SilentlyContinue).UninstallString
            if ($uninstallString) {
                # Use msiexec if possible
                if ($uninstallString -match "msiexec") {
                    # Extract product code GUID
                    if ($uninstallString -match "\{[0-9A-Fa-f\-]+\}") {
                        $guid = $matches[0]
                        Write-Host "Running msiexec uninstall for $guid"
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait
                    } else {
                        # fallback to running uninstall string directly
                        Write-Host "Running uninstall string: $uninstallString"
                        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /quiet /norestart" -Wait
                    }
                } else {
                    Write-Host "Running uninstall string: $uninstallString"
                    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /quiet /norestart" -Wait
                }
            } else {
                Write-Warning "No uninstall string found for $displayName"
            }
        }
    }
}

# 2. Remove Microsoft Store Python apps
Write-Host "Removing Microsoft Store Python apps..."
Get-AppxPackage *python* | Remove-AppxPackage -ErrorAction SilentlyContinue
Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like "*python*" | ForEach-Object {
    Write-Host "Removing provisioned package $($_.DisplayName)..."
    Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue
}

# 3. Delete common Python folders (adjust if you know more paths)
$foldersToRemove = @(
    "$env:LOCALAPPDATA\Programs\Python",
    "$env:ProgramFiles\Python*",
    "$env:ProgramFiles(x86)\Python*",
    "C:\Python*",
    "$env:APPDATA\Python",
    "$env:USERPROFILE\AppData\Local\Programs\Python*",
    "$env:USERPROFILE\AppData\Local\Microsoft\WindowsApps\python.exe",
    "$env:USERPROFILE\AppData\Local\Microsoft\WindowsApps\python3.exe"
)

foreach ($path in $foldersToRemove) {
    $items = Get-ChildItem -Path $path -Force -ErrorAction SilentlyContinue
    foreach ($item in $items) {
        Write-Host "Removing $($item.FullName)"
        Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 4. Search for python.exe and delete (dangerous, so restricted to common locations)
$searchPaths = @(
    "C:\Users\$env:USERNAME\AppData\Local\Programs",
    "C:\Program Files",
    "C:\Program Files (x86)",
    "C:\Python*"
)

foreach ($searchPath in $searchPaths) {
    Write-Host "Searching for python executables in $searchPath"
    $files = Get-ChildItem -Path $searchPath -Include python.exe, python3.exe -Recurse -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        Write-Host "Deleting $($file.FullName)"
        Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
    }
}

# 5. Clean Python from PATH environment variables
Write-Host "Removing Python entries from PATH environment variables..."
# User PATH
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -and $userPath -match "Python") {
    $newUserPath = ($userPath -split ';' | Where-Object { $_ -notmatch "Python" }) -join ';'
    [Environment]::SetEnvironmentVariable("PATH", $newUserPath, "User")
    Write-Host "Cleaned Python from User PATH"
}

# System PATH
$systemPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
if ($systemPath -and $systemPath -match "Python") {
    $newSystemPath = ($systemPath -split ';' | Where-Object { $_ -notmatch "Python" }) -join ';'
    [Environment]::SetEnvironmentVariable("PATH", $newSystemPath, "Machine")
    Write-Host "Cleaned Python from System PATH"
}

Write-Host "Complete! Restart your PC to finalize cleanup."
