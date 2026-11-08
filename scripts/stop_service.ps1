param (
    [int]$Port = 8501,
    [string]$ProjectDir = ""
)

if (-not $ProjectDir) {
    $ProjectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$checkScript = Join-Path $scriptDir "check_port_service.ps1"

$res = & powershell -NoProfile -ExecutionPolicy Bypass -File "$checkScript" -Port $Port -ProjectDir "$ProjectDir"

if ($res -like "*STATUS=PCGD_RUNNING*") {
    # Parse PID
    $tokens = $res -split ";"
    $pidToken = $tokens | Where-Object { $_ -like "PID=*" }
    $targetPid = ($pidToken -split "=")[1]
    
    Write-Host "[*] Tim thay tien trinh PCGD (PID: $targetPid). Dang dung..." -ForegroundColor Yellow
    taskkill /F /T /PID $targetPid | Out-Null
    Write-Host "[OK] Da dung sach se may chu PCGD (PID: $targetPid)!" -ForegroundColor Green
    exit 0
} elseif ($res -like "*STATUS=PORT_CONFLICT*") {
    Write-Host "[CANH BAO] Cong $Port dang bi chiem boi ung dung khac khong thuoc du an PCGD!" -ForegroundColor Red
    Write-Host "          He thong tu choi dung tien trinh nay de tranh lam anh huong ung dung khac." -ForegroundColor Yellow
    Write-Host "          Thong tin tien trinh: $res"
    exit 2
} else {
    Write-Host "[THONG BAO] He thong PCGD hien khong chay tren cong $Port." -ForegroundColor Cyan
    exit 0
}
