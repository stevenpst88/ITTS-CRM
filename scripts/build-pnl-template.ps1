# 重建 templates/pnl_template.xlsx（毛利分析 PNL 範本）
#
# 來源：公司原有的「報價單範本.xlsx」的「PNL (C)」頁籤（Profitability Calculation Worksheet for Implementation）。
# 原檔很髒（7 個外部連結、473 個已定義名稱、ActiveX、常駐註解、網路印表機設定），所以用 Excel COM 整理成乾淨的單頁範本，
# 再用 scripts/post-pnl-template.js 移除印表機設定與本機路徑。lib/quotePnlExcel.js 只負責「填值」，版面與公式都在這份範本裡。
#
# 用法（Windows，需安裝 Excel；PowerShell 5.1）：
#   powershell -ExecutionPolicy Bypass -File scripts\build-pnl-template.ps1 [-Source "<報價單範本.xlsx 路徑>"]
# 改範本的公式／版面後，要同步 lib/quotePnlExcel.js 的 REV_ROWS / COST_ROWS 與檔頭註解，並重跑 PNL 單元測試與 Excel 重算比對。
#
# 注意：本檔含中文，必須存成「UTF-8 with BOM」，PowerShell 5.1 才不會讀成亂碼。
param([string]$Source = "$env:USERPROFILE\OneDrive - 東捷資訊服務股份有限公司\Desktop\報價單範本.xlsx")
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$tmp  = Join-Path $env:TEMP 'pnl_template_build'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$work = Join-Path $tmp 'src_copy.xlsx'
$out  = Join-Path $tmp 'pnl_template_new.xlsx'
Copy-Item -LiteralPath $Source -Destination $work -Force
if (Test-Path $out) { Remove-Item $out -Force }

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false
try {
  $wb = $x.Workbooks.Open($work, 0, $false)     # UpdateLinks=0（不更新外部連結）
  $ws = $wb.Worksheets.Item('PNL (C)')

  # 1) 先切掉對「報價單 」工作表的跨表引用（F17、H17），之後才能刪掉那張表
  $ws.Range('F17').ClearContents() | Out-Null
  $ws.Range('H17').ClearContents() | Out-Null
  # 2) 刪除「報價單 」工作表；3) 斷開所有外部連結
  $wb.Worksheets.Item('報價單 ').Delete()
  $links = $wb.LinkSources(1)
  if ($links) { foreach ($l in $links) { $wb.BreakLink($l, 1) } }
  # 4) 刪除所有已定義名稱（只保留列印範圍 Print_Area）
  for ($i = $wb.Names.Count; $i -ge 1; $i--) {
    $n = $wb.Names.Item($i)
    if ($n.Name -notmatch '!Print_Area$') { $n.Delete() }
  }

  # 5) 公式修正
  $ws.Range('H19').Formula = '=SUM(H16:H18)'                          # 顧問收入合計：原本只加 H16:H17，但 16~18 都是輸入列
  $ws.Range('H59').Formula = '=F59*G59'; $ws.Range('H60').Formula = '=F60*G60'; $ws.Range('H61').Formula = '=SUM(H59:H60)'   # 軟體成本：原本 H61=G59（單價欄、不乘數量）
  $ws.Range('H64').Formula = '=F64*G64'; $ws.Range('H65').Formula = '=F65*G65'; $ws.Range('H66').Formula = '=SUM(H64:H65)'   # 硬體成本：原本加的是單價欄 G
  $ws.Range('H83').Formula = '=F83*G83'; $ws.Range('H84').Formula = '=F84*G84'                                              # 其他費用：數量×單價
  foreach ($r in 42..48) { $ws.Range("H$r").Formula = "=D$r*F$r" }                                                           # 隱藏列：統一成 Rates×Base Days
  $ws.Range('F103').Formula = '=H90+H85+H79'                                                                                 # 彙總表「其他」成本：原本漏掉差旅費 H79
  $ws.Range('H8').ClearContents() | Out-Null; $ws.Range('H9').ClearContents() | Out-Null                                    # 殘留的空白字串
  $ws.Range('F59').Value2 = 1                                                                                                  # 軟體成本數量預設 1

  # 6) 毛利率除零保護、顯示格式、移除常駐註解、改回一般檢視
  foreach ($c in 'C','D','E','F','G') { $ws.Range($c + '105').Formula = '=IF(' + $c + '102=0,0,' + $c + '104/' + $c + '102)' }   # 彙總表毛利%：某一類沒有收入時顯示 0.00%，不要 #DIV/0!
  $ws.Range('H93').Formula = '=IF(H13=0,0,H92/H13)'
  $ws.Range('F18').NumberFormat = $ws.Range('F17').NumberFormat                      # 第 3 列顧問收入的人天欄原本是貨幣格式（會顯示成 NT$3）
  $ws.Range('F83:F84').NumberFormat = $ws.Range('F17').NumberFormat                  # 其他費用數量：允許小數（中文版 Excel 不認 'General' 字串，所以沿用 F17 的格式）
  $ws.Range('G60').NumberFormat = $ws.Range('G59').NumberFormat
  $ws.Range('B33').NumberFormat = $ws.Range('B32').NumberFormat; $ws.Range('B33').Font.Name = $ws.Range('B32').Font.Name; $ws.Range('B33').Font.Size = $ws.Range('B32').Font.Size
  $ws.Range('B90').Value2 = '(由顧問主管依專案風險預估：0%、5%、10%、15%、20%)'   # 風險預留改由顧問主管選擇（原提示是「低-5%，中-10%，高-15%」）
  $ws.Range('A38:A39').HorizontalAlignment = -4131                                   # 顧問成本的項目說明改靠左（置中時超長文字會被左邊界切掉首字）
  while ($ws.Comments.Count -gt 0) { $ws.Comments.Item(1).Delete() | Out-Null }       # 原作者給填表人的提示註解（常駐顯示、蓋住小計標籤）
  $ws.Activate(); $x.ActiveWindow.View = 1                                            # 原檔存成分頁預覽（大字浮水印），改回一般檢視

  $x.CalculateFull()
  $ws.Range('A1').Select() | Out-Null
  $wb.SaveAs($out, 51)      # xlsx
  $wb.Close($false)
} finally { $x.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }

# 7) 後處理：移除網路印表機設定、本機路徑、作者姓名
node (Join-Path $repo 'scripts\post-pnl-template.js') $out (Join-Path $repo 'templates\pnl_template.xlsx')
Write-Output ('完成：' + (Join-Path $repo 'templates\pnl_template.xlsx'))
