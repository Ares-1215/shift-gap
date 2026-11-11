# 把班次表 JSON（export_day.py 產出的 .json.gz）、原始班次表（.docx／.xlsx）或加班需求清單（.csv／.xlsx）推到網站（Edge Function shiftgap）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File tools\push_to_site.ps1 -Path "C:\...\班次表_1106.json.gz" -Day 1106 -Name "運務通知_1106.docx"
#       powershell ... -Path "C:\...\加班預測_1106-1107.csv" -Kind needs
# 管理碼可放 tools\site.local.json（不入庫）：{ "admin": "....", "user": "自動排程" }
# 同一天再推一次就是覆蓋；清單的日期由 CSV「日期」欄算出（例 1106-1107）
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [ValidateSet('schedule', 'needs', 'matrix', 'actual', 'fleet')][string]$Kind = 'schedule',
  [string]$Day,
  [string]$Name,
  [string]$Admin,
  [string]$User
)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$FN = 'https://hmqnlovyzlvvnkqmfwtt.supabase.co/functions/v1/shiftgap'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$local = Join-Path $here 'site.local.json'
if ((-not $Admin -or -not $User) -and (Test-Path $local)) {
  $c = Get-Content $local -Raw -Encoding UTF8 | ConvertFrom-Json
  if (-not $Admin) { $Admin = [string]$c.admin }
  if (-not $User)  { $User  = [string]$c.user }
}
if (-not $Admin) { throw '缺管理碼：用 -Admin 指定，或放在 tools\site.local.json' }
if (-not $User)  { $User = '自動排程' }

$f = Get-Item -LiteralPath $Path
$name = if ($Name) { $Name } else { $f.Name }
$mime = switch -Regex ($f.Name) {
  '\.json\.gz$' { 'application/json+gzip'; break }
  '\.json$'     { 'application/json'; break }
  '\.xlsx$'     { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; break }
  '\.docx$'     { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'; break }
  '\.csv$'      { 'text/csv'; break }
  default       { 'application/octet-stream' }
}
$summary = $null
if ($Kind -eq 'needs' -and $f.Extension -ieq '.csv') {
  $rows = @(Import-Csv -LiteralPath $f.FullName -Encoding UTF8)
  $days = @($rows | ForEach-Object { [string]$_.'日期' -replace '\D', '' } | Where-Object { $_ } |
            ForEach-Object { $_.Substring([Math]::Max(0, $_.Length - 4)) } | Sort-Object -Unique)
  $summary = @{ days = $days; rows = $rows.Count }
  if (-not $Day -and $days.Count) {
    if ($days.Count -le 4) { $Day = ($days -join '-') } else { $Day = $days[0] + '-' + $days[-1] }
  }
}
if (-not $Day) {
  $m = [regex]::Match($name, '(1[0-2]|0?[1-9])(0[1-9]|[12]\d|3[01])(?!\d)')
  if ($m.Success) { $Day = $m.Value } else { throw "檔名「$name」沒有日期（MMDD），請用 -Day 指定" }
}

$b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($f.FullName))
$payload = @{ action = 'put'; code = $Admin; kind = $Kind; day = $Day; name = $name; mime = $mime; data = $b64; uploader = $User }
if ($summary) { $payload.summary = $summary }
$json = ConvertTo-Json -InputObject $payload -Depth 5 -Compress
$r = Invoke-RestMethod -Uri $FN -Method Post -ContentType 'application/json; charset=utf-8' `
  -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 180
Write-Host ("已推到網站：{0} → {1}（{2} KB，{3}）" -f $name, $r.key, [math]::Round($r.size / 1KB), $r.uploaded_at)
