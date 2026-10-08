# 重建 templates/pnl_template.xlsx（毛利分析 PNL 範本，供成本明細 costLines 使用）
#
# ★ 來源檔需自備、且不可放進 repo：來源是公司最新的 PNL 範本活頁簿（含「PNL 」工作表，名稱尾端有空白，
#   另有一張隱藏的「(8+1)」報價單工作表），內含客戶資料（客戶名、人名、廠商名、真實費率）。
#   本腳本不寫死任何路徑，來源檔用 -Source 傳入；請勿 commit 來源檔，也不要 commit 含客戶資料的中間檔（%TEMP%\pnl_template_build 內）。
# 範本特色：交際費用、印花稅、差旅各列、Free Lancer(委外) 欄等成本明細區（對應 q.costLines 的五個分類）。
# 本腳本用 Excel COM 清空範例資料、修公式、擴充資料列容量，再用 scripts/post-pnl-template.js 做 JSZip 後處理
# （移除印表機設定與本機路徑、展開共用公式、移除 calcChain、設 fullCalcOnLoad）。
# lib/quotePnlExcel.js 只負責「填值」，版面與公式都在這份範本裡；填值的儲存格位置見該檔的 LAYOUT（改版面時兩邊要同步）。
#
# 用法（Windows，需安裝 Excel；PowerShell 5.1）：
#   powershell -ExecutionPolicy Bypass -File scripts\build-pnl-template.ps1 -Source "<範本.xlsx 路徑>" [-SheetName PNL] [-Out "<輸出路徑>"]
#   -SheetName：要保留的工作表名稱（比對時會 Trim 掉前後空白）；其餘工作表一律刪除。
#   -Out：預設 templates\pnl_template.xlsx。
#
# 擴充後的列號（以原範本為基準 +12 列）：軟體資料列 68-73（原 68-69）、硬體 77-82（原 73）、差旅 87-94（原 78-85）、
# 其他費用 99-104（原 90-92）；顧問成本 37-49 / 51-57 不變。每區 Total 的 SUM 與每個資料列的 H=F*G 由本腳本明確寫入。
#
# 注意：本檔含中文，必須存成「UTF-8 with BOM」，PowerShell 5.1 才不會讀成亂碼。
param(
  [Parameter(Mandatory = $true)][string]$Source,
  [string]$SheetName = 'PNL',
  [string]$Out = ''
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
if (-not $Out) { $Out = Join-Path $repo 'templates\pnl_template.xlsx' }
$tmp  = Join-Path $env:TEMP 'pnl_template_build'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$work = Join-Path $tmp 'src_copy.xlsx'
$com  = Join-Path $tmp 'pnl_com.xlsx'     # Excel 存出、尚未後處理的版本（共用公式還在）
Copy-Item -LiteralPath $Source -Destination $work -Force
if (Test-Path $com) { Remove-Item $com -Force }

# 擴充列數（資料列總數 = 原本 + 擴充）：軟體 2→6、硬體 1→6、其他費用 3→6
$AddSw = 4; $AddHw = 5; $AddOt = 3
$Shift = $AddSw + $AddHw + $AddOt      # 全部擴充列數 = 12；原範本第 r 列（r 在其他費用 Total 之後）→ r + 12

function Assert-Label($ws, $addr, $pattern) {
  $t = [string]$ws.Range($addr).Text
  if ($t -notmatch $pattern) { throw "範本版面與預期不符：$addr = [$t]，預期符合 $pattern（來源範本改版了？請先檢查再重跑）" }
}

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false
try {
  $wb = $x.Workbooks.Open($work, 0, $false)     # UpdateLinks=0（不更新外部連結）
  $ws = $null
  foreach ($s in $wb.Worksheets) { if (([string]$s.Name).Trim() -ieq $SheetName) { $ws = $s } }
  if (-not $ws) { throw "來源活頁簿找不到工作表 [$SheetName]" }

  # 0) 版面斷言：確認是預期的新範本（以原列號）
  Assert-Label $ws 'A36'  'Consultant Role'
  Assert-Label $ws 'G70'  'Total Cost'
  Assert-Label $ws 'G74'  'Total Cost'
  Assert-Label $ws 'G86'  'Total Cost'
  Assert-Label $ws 'G93'  'Total Cost'
  Assert-Label $ws 'A91'  '0\.001'
  if ([string]$ws.Range('G91').Formula -ne '=H13*0.001') { throw 'G91 公式不是 =H13*0.001' }

  # 1) 其他工作表若被 PNL 引用就中止（新範本實測沒有跨表引用）
  $others = @(); foreach ($s in $wb.Worksheets) { if ($s.Name -ne $ws.Name) { $others += $s.Name } }
  foreach ($c in $ws.UsedRange.Cells) {
    if ($c.HasFormula) {
      $f = [string]$c.Formula
      if ($f -match '\[') { throw ('外部活頁簿引用：' + $c.Address($false, $false) + ' ' + $f) }
      foreach ($on in $others) { if ($f.Contains($on + '!') -or $f.Contains("'" + $on + "'!")) { throw ('跨表引用：' + $c.Address($false, $false) + ' ' + $f) } }
    }
  }
  # 2) 刪除其他工作表（含隱藏的「(8+1)」報價單）；3) 斷開外部連結；4) 刪除所有已定義名稱（列印範圍最後重設）
  foreach ($n in $others) { $wb.Worksheets.Item($n).Delete() | Out-Null }
  $links = $wb.LinkSources(1)
  if ($links) { foreach ($l in $links) { $wb.BreakLink($l, 1) } }
  for ($i = $wb.Names.Count; $i -ge 1; $i--) { try { $wb.Names.Item($i).Delete() } catch { } }
  $ws.Name = 'PNL'                                  # 去掉名稱尾端的空白

  # 5) 移除原作者留給填表人的註解（含人名與真實費率提示）
  while ($ws.Comments.Count -gt 0) { $ws.Comments.Item(1).Delete() | Out-Null }

  # 6) 清空範例／敏感資料（只清內容，格式與公式保留）。以下皆為「原範本」列號（擴充列之前）
  foreach ($a in 'B5','H5','B6','B7','B8','B9','B10','G10','H10','B11','H11','H7','H8','H9',   # 抬頭：申請人、日期、客戶、PM、專案…
                 'H14','H27','H31',                                                            # 寫死的收入
                 'B16:H18','B23:H24','B28:G28','B32:H32',                                      # 收入明細列（日期／百分比／金額）
                 'A37:D49','F37:F49','A51:D57','F51:F57',                                      # 顧問成本：角色、委外廠商、費率、基本天數
                 'A78:B85','D78:D85','F78:G85') {                                              # 差旅範例列
    try { $rg = $ws.Range($a); if ($rg.Cells.Count -eq 1) { $rg = $rg.MergeArea }; $rg.ClearContents() | Out-Null } catch { throw "清除 $a 失敗：$($_.Exception.Message)" }
  }

  # 7) 容量擴充（由下往上插入，列號才不會互相影響）。新列格式複製上一列（資料列），公式另外明確寫入
  Assert-Label $ws 'G93' 'Total Cost'
  $ws.Range('93:' + (93 + $AddOt - 1)).EntireRow.Insert(-4121, 0) | Out-Null    # 其他費用：Total 列(93) 之前插入
  Assert-Label $ws 'G74' 'Total Cost'
  $ws.Range('74:' + (74 + $AddHw - 1)).EntireRow.Insert(-4121, 0) | Out-Null    # 硬體
  Assert-Label $ws 'G70' 'Total Cost'
  $ws.Range('70:' + (70 + $AddSw - 1)).EntireRow.Insert(-4121, 0) | Out-Null    # 軟體

  # 擴充後的列號
  $sw1 = 68;  $sw2 = 68 + 1 + $AddSw;  $swT = $sw2 + 1          # 68-73, Total 74
  $hw1 = $swT + 3; $hw2 = $hw1 + $AddHw; $hwT = $hw2 + 1        # 77-82, Total 83
  $tr1 = $hwT + 4; $tr2 = $tr1 + 7; $trT = $tr2 + 1             # 87-94, Total 95
  $ot1 = $trT + 4; $ot2 = $ot1 + 2 + $AddOt; $otT = $ot2 + 1    # 99-104, Total 105
  Assert-Label $ws "G$swT" 'Total Cost'
  Assert-Label $ws "G$hwT" 'Total Cost'
  Assert-Label $ws "G$trT" 'Total Cost'
  Assert-Label $ws "G$otT" 'Total Cost'
  if ($ot2 -ne 104 -or $swT -ne 74 -or $hwT -ne 83 -or $trT -ne 95 -or $otT -ne 105) { throw '擴充後列號與預期不符' }

  # 8) 公式修正（逐項列在 lib/quotePnlExcel.js 檔尾「範本與原檔的差異」）
  $ws.Range('H13').Formula = '=H14+H22+H27+H31'                                   # 總收入：原本漏掉 4.其他收入 H31（QC 欄 H122 會不平）
  foreach ($r in (37..49) + (51..57)) {                                           # 顧問成本：每列統一 = Rates × Base Days × Months/Year
    $ws.Range("G$r").Value2 = 1                                                   #   原本有的列沒有 G（H=D*F*G 恆為 0）、有的列公式漏乘 G；統一預設 1
    $ws.Range("H$r").Formula = "=D$r*F$r*G$r"
  }
  $ws.Range('F50').Formula = '=SUM(F37:F49)'; $ws.Range('H50').Formula = '=SUM(H37:H49)'   # Subtotal(1)：原本 SUM 到 48，漏掉第 49 列（同樣是資料列）
  foreach ($rr in @(@($sw1, $sw2, $swT), @($hw1, $hw2, $hwT), @($tr1, $tr2, $trT), @($ot1, $ot2, $otT))) {
    foreach ($r in $rr[0]..$rr[1]) { $ws.Range("H$r").Formula = "=F$r*G$r" }      # 軟體／硬體／差旅／其他：數量×單價（原本軟硬體與差旅多數列是空的、要手填金額）
    $ws.Range('H' + $rr[2]).Formula = '=SUM(H' + $rr[0] + ':H' + $rr[1] + ')'     # 各區 Total 明確涵蓋所有資料列
  }
  $rMargin = 100 + $Shift; $rPct = 101 + $Shift                                           # 112、113
  $ws.Range("H$rPct").Formula = "=IF(H13=0,0,H$rMargin/H13)"                      # 總毛利率：除零保護（資料清空時不顯示 #DIV/0!）
  $rSumRev = 110 + $Shift; $rSumCost = 111 + $Shift; $rSumGp = 112 + $Shift; $rSumPct = 113 + $Shift   # 122、123、124、125
  foreach ($c in 'C','D','E','F','G') {
    $ws.Range($c + $rSumPct).Formula = '=IF(' + $c + $rSumRev + '=0,0,' + $c + $rSumGp + '/' + $c + $rSumRev + ')'   # 彙總表毛利%：原本只有 C、G 且沒除零保護
  }
  $ws.Range('D' + $rSumPct + ':F' + $rSumPct).NumberFormat = $ws.Range('C' + $rSumPct).NumberFormat

  # 8b) 範例列清空後字型會退回儲存格底層字型（原本是綠色的範例文字）；差旅第 2~4 列 A、D 欄統一成第 1 列的字型，避免之後系統填入的文字顏色不一致
  foreach ($col in 'A', 'D') {
    $ref = $ws.Range("$col$tr1")
    foreach ($r in ($tr1 + 1)..($tr1 + 3)) { $d = $ws.Range("$col$r"); $d.Font.Color = $ref.Font.Color; $d.Font.Size = $ref.Font.Size; $d.Font.Name = $ref.Font.Name }
  }

  # 8c) 原範本這幾格是 General（填入數字會顯示成沒有 NT$/千分位、對齊也不同）→ 沿用同區其他列的格式（舊版範本也做過同樣處理）
  $ws.Range("G${sw1}:G$sw2").NumberFormat = $ws.Range("G$hw1").NumberFormat                         # 軟體單價
  $ws.Range("F$($tr2 - 1):F$tr2").HorizontalAlignment = $ws.Range("F$tr1").HorizontalAlignment      # 差旅第 7、8 列數量的對齊（格式見 8c2）
  $ws.Range("G$tr2").NumberFormat = $ws.Range("G$tr1").NumberFormat                                 # 差旅第 8 列單價

  # 8c2) 「數量/次數」F 欄：差旅 87-94、其他費用 99、101-104 原本是整數格式 #,##0（會把 1.5 顯示成 2、0.1 顯示成 0，值是對的只是顯示被四捨五入）。
  #      改成 General，與顧問(37-57)、軟體(68-73)、硬體(77-82) 的 F 欄一致：1.5 顯示 1.5、0.1 顯示 0.1、10 顯示 10，網頁預覽器 (lib/xlsxSheetHtml.js) 的顯示也一致。
  #      不用 #,##0.##：整數會在 Excel 顯示成「5.」（尾端多一個小數點）；代價是 ≥1000 的數量沒有千分位（數量/次數很少到這個量級）。
  #      印花稅列 F100（固定數量 1，格式 0.000 → 顯示 1.000）不在這次範圍，維持原樣。
  #      ★ 不能寫 NumberFormat = 'General'：繁體中文版 Excel 的 COM 會拒絕（「無法設定種類 Range 的 NumberFormat 屬性」，格式字串要寫 'G/通用格式'），
  #        所以從原本就是 General 的顧問數量格 F37 讀格式字串再指定（跨語系都可用；寫出的 xlsx 是 numFmtId=0 的 General）。
  $fmtGeneral = [string]$ws.Range('F37').NumberFormat
  $ws.Range("F${tr1}:F$tr2").NumberFormat = $fmtGeneral                                               # 差旅 87-94
  $ws.Range("F$ot1").NumberFormat = $fmtGeneral                                                      # 其他費用 99（交際費）
  $ws.Range("F$($ot1 + 2):F$ot2").NumberFormat = $fmtGeneral                                         # 其他費用 101-104（100 是印花稅列）

  # 8d) F 欄（彙總表「其他」成本）原寬 11.8，成本 ≥ 約 100 萬（含風險預留）時會顯示 #######；加寬成與 E 欄（硬體）相同
  $ws.Columns.Item(6).ColumnWidth = $ws.Columns.Item(5).ColumnWidth

  # 9) 錯字與提示文字
  $ws.Range('A22').Value2 = ([string]$ws.Range('A22').Value2).Replace('Licnese', 'License')                                 # Licnese → License（兩處）
  $rLic = 66
  $ws.Range("A$rLic").Value2 = ([string]$ws.Range("A$rLic").Value2).Replace('Licnese', 'License')
  $rRisk = 95 + $Shift
  $ws.Range("B$rRisk").Value2 = ([string]$ws.Range("B$rRisk").Value2).Replace('Risk leve', 'Risk level')                    # Project Risk leve → level
  $rSig = 107 + $Shift
  $ws.Range("A$rSig").Value2 = ([string]$ws.Range("A$rSig").Value2).Replace('Consultin ', 'Consultant ')                    # Consultin Director → Consultant Director
  $ws.Range('J31').Value2 = ([string]$ws.Range('J31').Value2).Replace('請其他', '請填其他')                                   # 「請其他收入金額」→「請填其他收入金額」
  $ws.Range('B' + (98 + $Shift)).Value2 = '(由顧問主管依專案風險預估：0%、5%、10%、15%、20%)'   # 風險預留改由顧問主管選擇（原提示是「低-5%，中-10%，高-15%」；與舊範本一致）

  # 10) 檢視與列印：一般檢視、回到左上角、列印範圍＝PNL 自己的 A1:H(114+12)（含底部彙總表）
  $ws.Activate()
  $x.ActiveWindow.View = 1
  $x.ActiveWindow.Zoom = 120
  $x.ActiveWindow.ScrollRow = 1; $x.ActiveWindow.ScrollColumn = 1
  $ws.PageSetup.PrintArea = '$A$1:$H$' + (114 + $Shift)

  $x.CalculateFull()
  # 錯誤值掃描（清空後不應有任何錯誤）
  $errs = @()
  foreach ($c in $ws.UsedRange.Cells) { if ($c.HasFormula) { $t = [string]$c.Text; if ($t -match '^#') { $errs += ($c.Address($false, $false) + '=' + $t) } } }
  if ($errs.Count -gt 0) { throw ('範本含錯誤值：' + ($errs -join ', ')) }
  $ws.Range('A1').Select() | Out-Null
  $wb.SaveAs($com, 51)      # xlsx
  $wb.Close($false)
} finally { $x.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }

# 11) 後處理：移除網路印表機設定、本機路徑、作者姓名；展開共用公式；移除 calcChain；設 fullCalcOnLoad
node (Join-Path $repo 'scripts\post-pnl-template.js') $com $Out
Write-Output ('完成：' + $Out)
