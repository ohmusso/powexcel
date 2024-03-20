Import-Module "${PSScriptRoot}\../excel.psm1" -Force

# test.xlsx����f�[�^��ǂݏo��
$excel = Get-Excel -startDir "C:\Users\toshitaka\source\repos\powershell\powexcel\test"
$sheet = $sheet = $excel.ActiveWorkbook.Sheets("Sheet1")
$objs = Read-Table -startCell "B2" -rowOffset 0 -colOffset 0 -headerRow 1 -sheet $sheet

# close excel
$sheet = $null
$excel.Quit()
$excel = $null

# �V�K�쐬
New-Excel -filePath "${PSScriptRoot}\book1.xlsx" -startCell "B2" -objects $objs
