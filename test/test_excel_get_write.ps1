Import-Module "${PSScriptRoot}\../excel.psm1"

# test.xlsx����f�[�^��ǂݏo��
$excel = Get-Excel -startDir "C:\Users\toshitaka\source\repos\powershell\powexcel\test"
$sheet = $sheet = $excel.ActiveWorkbook.Sheets("Sheet1")
$objs = Read-Table -startCell "B2" -rowOffset 0 -colOffset 0 -headerRow 1 -sheet $sheet

# ��������
Write-Table "F2" $sheet $objs $true

# close excel
$sheet = $null
$excel.Quit()
$excel = $null
