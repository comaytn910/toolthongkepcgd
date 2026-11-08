param (
    [int]$Port = 8501,
    [string]$ProjectDir = ""
)

if (-not $ProjectDir) {
    $ProjectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

# Chuan hoa duong dan du an
$normProjectDir = [System.IO.Path]::GetFullPath($ProjectDir).ToLower()

# Tim PID dang lang nghe tren cong
$connection = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $connection) {
    Write-Output "STATUS=NOT_RUNNING"
    exit 0
}

$pidListening = $connection.OwningProcess

try {
    $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $pidListening" -ErrorAction Stop
    $cmd = $proc.CommandLine
    $procName = $proc.Name
    
    # Kiem tra xem co phai tien trinh PCGD cua du an khong
    $isPCGD = $false
    if ($cmd -and ($cmd -like "*streamlit*" -and ($cmd -like "*ui/app.py*" -or $cmd -like "*ui\app.py*"))) {
        $isPCGD = $true
    }
    
    if ($isPCGD) {
        Write-Output "STATUS=PCGD_RUNNING;PID=$pidListening;NAME=$procName"
    } else {
        Write-Output "STATUS=PORT_CONFLICT;PID=$pidListening;NAME=$procName;CMD=$cmd"
    }
} catch {
    Write-Output "STATUS=PORT_CONFLICT;PID=$pidListening;NAME=UNKNOWN;CMD="
}
