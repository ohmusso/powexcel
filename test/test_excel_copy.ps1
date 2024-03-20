Import-Module "${PSScriptRoot}\../excel.psm1" -Force

# 予めコピー先のエクセル test_copy_to.xlsx を開いておく
$excel = Get-Excel -isOpen $false
# コピー元のシート
$fromBook = $excel.Workbooks.Open("${PSScriptRoot}\test_copy_from.xlsx")
$fromSheet = $fromBook.Sheets("Sheet1")
# コピー先のシート
$destBook = $excel.Workbooks("test_copy_to.xlsx")
$destSheet = $destBook.Sheets("Sheet1")

# コピー
Copy-Table -fromSheet $fromSheet -fromRangeStr "B2:G2" -destSheet $destSheet -destCellStr "A2"

# コピー元は保存せずに閉じる
$fromBook.Close($false)
