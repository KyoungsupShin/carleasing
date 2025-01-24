# 모든 Excel 프로세스를 찾고 종료합니다.
Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }

# 종료된 프로세스를 확인합니다.
if (Get-Process excel -ErrorAction SilentlyContinue) {
    Write-Output "Excel 프로세스를 종료하지 못했습니다."
} else {
    Write-Output "모든 Excel 프로세스가 성공적으로 종료되었습니다."
}
