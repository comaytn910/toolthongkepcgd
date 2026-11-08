param (
    [string]$ProjectDir = ""
)

if (-not $ProjectDir) {
    $ProjectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

$WshShell = New-Object -ComObject WScript.Shell
$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$TargetBat = Join-Path $ProjectDir "KHOI_DONG_HE_THONG.bat"
$WorkDir = $ProjectDir

# Xoa cac file shortcut cu
Get-ChildItem -Path $DesktopPath -Filter "*PCGD*.lnk" | ForEach-Object {
    if ($_.Name -ne "HE THONG PCGD 2026 - THAI NGUYEN.lnk") {
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
    }
}

# Tao loi tat chuan
$ShortcutPath = Join-Path $DesktopPath "HE THONG PCGD 2026 - THAI NGUYEN.lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetBat
$Shortcut.WorkingDirectory = $WorkDir
$Shortcut.Description = "He thong Tong hop so lieu Pho cap Giao duc va Xoa mu chu nam 2026 - So GDDT Thai Nguyen"
$Shortcut.IconLocation = "$env:SystemRoot\System32\imageres.dll,114"
$Shortcut.Save()

Write-Host "SUCCESS: $ShortcutPath"
