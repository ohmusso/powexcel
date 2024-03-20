Import-Module "${PSScriptRoot}\../excel.psm1" -Force

# �\�߃R�s�[��̃G�N�Z�� test_copy_to.xlsx ���J���Ă���
$excel = Get-Excel -isOpen $false
# �R�s�[���̃V�[�g
$fromBook = $excel.Workbooks.Open("${PSScriptRoot}\test_copy_from.xlsx")
$fromSheet = $fromBook.Sheets("Sheet1")
# �R�s�[��̃V�[�g
$destBook = $excel.Workbooks("test_copy_to.xlsx")
$destSheet = $destBook.Sheets("Sheet1")

# �R�s�[
Copy-Table -fromSheet $fromSheet -fromRangeStr "B2:G2" -destSheet $destSheet -destCellStr "A2"

# �R�s�[���͕ۑ������ɕ���
$fromBook.Close($false)
