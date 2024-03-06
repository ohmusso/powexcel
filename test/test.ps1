Import-Module "${PSScriptRoot}\../excel.psm1"

# test.xlsx����f�[�^��ǂݏo��
$excel = Get-Excel -startDir "C:\Users\toshitaka\source\repos\powershell\powexcel\test"
$sheet = $sheet = $excel.ActiveWorkbook.Sheets("Sheet1")
Read-Table -startCell "B2" -rowOffset 0 -colOffset 0 -headerRow 2 -sheet $sheet

# close excel
$sheet = $null
$excel.Quit()
$excel = $null