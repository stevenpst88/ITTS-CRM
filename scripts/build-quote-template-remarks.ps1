# 報價單範本（templates/quotation_template.xlsx）Remarks 區改版：讓付款方式與追加條款可以「活的」
#
# 做了什麼（只在範本上做一次，已套用；重跑會被保護檢查擋下）：
#   1) 第 3 條「付款方式」那一列（第 86 列）合併 B:J、自動換行，長句（分 5 期）才放得下
#   2) 在第 6 條之後插入 8 列（第 90~97 列）＝追加條款（第 7~14 條）用，每列合併 B:J、自動換行、預設「隱藏」；
#      輸出時 lib/quoteExcel.js 依條款數展開（LAYOUT.extraFirst/extraLast）。原本的空白列順延到第 98 列（維持條款與簽名區之間的間距）
#   3) 簽名區（Customer Confirme by／Prepared by）整體下移 8 列：LAYOUT.signCell G92→G100、報價專用章錨點 row 91→99、列印範圍 A1:K95→A1:K103
# 之後再用 scripts/post-quote-template.js 移除印表機設定／本機路徑／calcChain。
#
# 用法（Windows＋Excel，PowerShell 5.1）：powershell -ExecutionPolicy Bypass -File scripts\build-quote-template-remarks.ps1 [-Out <輸出路徑>]
# 注意：本檔含中文，必須存成 UTF-8 with BOM。
param([string]$Out = '')
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $repo 'templates\quotation_template.xlsx'
$tmp  = Join-Path $env:TEMP 'quote_template_build'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$work = Join-Path $tmp 'work.xlsx'
$raw  = Join-Path $tmp 'saved.xlsx'
if (-not $Out) { $Out = $src }
Copy-Item -LiteralPath $src -Destination $work -Force
if (Test-Path $raw) { Remove-Item $raw -Force }

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Open($work, 0, $false)
  $ws = $wb.Worksheets.Item(1)
  if ([string]$ws.Range('B91').Text -like 'Customer Confirme*' -eq $false -and [string]$ws.Range('B99').Text -like 'Customer Confirme*') { throw '範本已經套用過（Customer Confirme 已在第 99 列）；不要重複執行。' }
  if (-not ([string]$ws.Range('B91').Text -like 'Customer Confirme*')) { throw '範本版面不是預期的（第 91 列應為 Customer Confirme by）。' }

  $ref = $ws.Range('B84')
  function Format-Remark($rng) {
    $rng.Font.Name = $ref.Font.Name; $rng.Font.Size = $ref.Font.Size; $rng.Font.Bold = $ref.Font.Bold; $rng.Font.Color = $ref.Font.Color
    $rng.Merge()
    $rng.WrapText = $true; $rng.VerticalAlignment = -4160; $rng.HorizontalAlignment = -4131   # 靠上、靠左
  }
  # 1) 第 3 條付款方式：合併＋換行
  Format-Remark $ws.Range('B86:J86')
  # 2) 插入 8 列（追加條款）
  $ws.Rows('90:97').Insert(-4121) | Out-Null                # xlShiftDown
  foreach ($r in 90..97) {
    $ws.Range("B${r}:J${r}").UnMerge() | Out-Null
    Format-Remark $ws.Range("B${r}:J${r}")
    $ws.Rows($r).RowHeight = 18
    $ws.Rows($r).Hidden = $true
  }
  $ws.Activate(); $ws.Range('A1').Select() | Out-Null
  $x.CalculateFull()
  $wb.SaveAs($raw, 51)
  $wb.Close($false)
} finally { $x.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }

node (Join-Path $repo 'scripts\post-quote-template.js') $raw $Out
Write-Output ('完成：' + $Out)
