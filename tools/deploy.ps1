# 班次空檔分析：建表 + 部署 Edge Function shiftgap（走 Supabase Management API，不需 supabase CLI）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File tools\deploy.ps1 [-SkipSql] [-SkipDeploy]
# 需要：notebookLM\.mcp.json 的 mcpServers.supabase.env.SUPABASE_ACCESS_TOKEN（PAT）
#       PAT 失效時到 https://supabase.com/dashboard/account/tokens 產生新的，貼回 .mcp.json 再跑
param([switch]$SkipSql, [switch]$SkipDeploy)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$ref  = 'hmqnlovyzlvvnkqmfwtt'
$slug = 'shiftgap'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$src  = Join-Path $root "supabase\functions\$slug\index.ts"
$sql  = Join-Path $root "supabase\migrations\sg_files.sql"
foreach ($p in @($src, $sql)) { if (-not (Test-Path $p)) { throw "找不到檔案：$p" } }

$cfg = Get-Content (Join-Path $env:USERPROFILE 'notebookLM\.mcp.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$pat = $cfg.mcpServers.supabase.env.SUPABASE_ACCESS_TOKEN
if (-not $pat) { throw '.mcp.json 裡找不到 mcpServers.supabase.env.SUPABASE_ACCESS_TOKEN' }
$h = @{ Authorization = "Bearer $pat" }

Write-Host '1/4 檢查 PAT…'
try { $null = Invoke-RestMethod -Uri "https://api.supabase.com/v1/projects/$ref/functions" -Headers $h -TimeoutSec 30 }
catch { throw "PAT 失效或無權限：$($_.Exception.Message)`n→ 到 https://supabase.com/dashboard/account/tokens 產生新 PAT，更新 notebookLM\.mcp.json 後重跑本腳本。" }

if (-not $SkipSql) {
  Write-Host '2/4 建表 sg_files / sg_settings（已存在就跳過，通行碼不會被覆蓋）…'
  # 用 .NET 讀成純字串：Get-Content 的字串帶 PSPath 等屬性，ConvertTo-Json 會把它變成物件（API 回 expected string, received object）
  $q = [IO.File]::ReadAllText($sql, [Text.Encoding]::UTF8)
  $body = ConvertTo-Json -InputObject @{ query = $q } -Compress
  $null = Invoke-RestMethod -Uri "https://api.supabase.com/v1/projects/$ref/database/query" -Headers $h -Method Post `
    -ContentType 'application/json' -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 60
  Write-Host '   建表完成。'
  # 通行碼：從 tools\site.local.json（不入庫）寫進 sg_settings；沒有就留佔位值，之後自己 update
  $local = Join-Path $here 'site.local.json'
  if (Test-Path $local) {
    $c = Get-Content $local -Raw -Encoding UTF8 | ConvertFrom-Json
    $sets = @()
    if ($c.view)  { $sets += "update public.sg_settings set value = '$(([string]$c.view) -replace "'", "''")' where key = 'view_code';" }
    if ($c.admin) { $sets += "update public.sg_settings set value = '$(([string]$c.admin) -replace "'", "''")' where key = 'admin_code';" }
    if ($sets.Count) {
      $body2 = ConvertTo-Json -InputObject @{ query = ($sets -join "`n") } -Compress
      $null = Invoke-RestMethod -Uri "https://api.supabase.com/v1/projects/$ref/database/query" -Headers $h -Method Post `
        -ContentType 'application/json' -Body ([Text.Encoding]::UTF8.GetBytes($body2)) -TimeoutSec 60
      Write-Host "   通行碼已從 site.local.json 寫入（$($sets.Count) 組）。"
    }
  } else { Write-Warning '   沒有 tools\site.local.json，通行碼仍是佔位值；請加 { "view": "...", "admin": "..." } 後重跑，或自己 update sg_settings。' }
} else { Write-Host '2/4 略過建表' }

if (-not $SkipDeploy) {
  Write-Host "3/4 部署 Edge Function $slug（verify_jwt=false，函式自己驗通行碼）…"
  Add-Type -AssemblyName System.Net.Http
  $client = New-Object System.Net.Http.HttpClient
  $client.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $pat)
  $form = New-Object System.Net.Http.MultipartFormDataContent
  $meta = '{"entrypoint_path":"index.ts","name":"' + $slug + '","verify_jwt":false}'
  $form.Add((New-Object System.Net.Http.StringContent($meta, [Text.Encoding]::UTF8, 'application/json')), 'metadata')
  $bytes = [IO.File]::ReadAllBytes($src)
  $fileContent = New-Object System.Net.Http.ByteArrayContent(,$bytes)
  $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('application/typescript')
  $form.Add($fileContent, 'file', 'index.ts')
  $resp = $client.PostAsync("https://api.supabase.com/v1/projects/$ref/functions/deploy?slug=$slug", $form).Result
  $txt  = $resp.Content.ReadAsStringAsync().Result
  if (-not $resp.IsSuccessStatusCode) { throw "部署失敗 $([int]$resp.StatusCode)：$txt" }
  Write-Host '   部署完成。'
} else { Write-Host '3/4 略過部署' }

Write-Host '4/4 測試呼叫（錯誤通行碼應回 401）…'
Start-Sleep -Seconds 3
$code = 0
try {
  $null = Invoke-WebRequest -Uri "https://$ref.supabase.co/functions/v1/$slug" -Method Post -ContentType 'application/json' `
    -Body '{"action":"login","code":"wrong"}' -UseBasicParsing -TimeoutSec 30
  $code = 200
} catch { if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode } else { throw } }
if ($code -eq 401) { Write-Host '   OK：函式上線，通行碼驗證正常。' } else { Write-Warning "   函式回應 $code，請檢查。" }
