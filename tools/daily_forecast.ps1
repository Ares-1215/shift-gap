# 班次空檔分析：每天自動推測隔日加班需求清單（＋有通知就配車）
# 由 Windows 工作排程「ShiftGap-Forecast」每天 13:40 呼叫（OpsHub-Daily 13:00 同步 205 之後）；漏跑會在下次開機補執行
$env:PYTHONIOENCODING = 'utf-8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$tools = Split-Path -Parent $MyInvocation.MyCommand.Path
$log = Join-Path $tools 'daily.log'
try {
  & "C:\Users\26516\AppData\Local\Programs\Python\Python312\python.exe" (Join-Path $tools 'daily_forecast.py') 2>&1 | Out-Null
} catch {
  "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] 錯誤: $_" | Add-Content -Encoding utf8 $log
}
if ((Test-Path $log) -and (Get-Item $log).Length -gt 300KB) {
  $tail = Get-Content $log -Tail 300
  Set-Content -Encoding utf8 $log $tail
}
