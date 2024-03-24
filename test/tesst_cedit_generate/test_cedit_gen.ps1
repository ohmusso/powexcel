using namespace System.Text

Import-Module "${PSScriptRoot}\../../cedit.psm1" -Force

$code = [StringBuilder]::new()
$funcBody = [StringBuilder]::new()

Add-CodeLine -code $code -line "uint8 ucData = 0;" -comment "comment"
Add-CodeLine -code $code -line "uint16 usData = 0;" -comment "comment"

Add-CodeLine -code $funcBody -line "ucData = 0;" -comment "ucData" -indent 1
Add-CodeLine -code $funcBody -line "usData = 0;" -comment "usData" -indent 1
Add-CodeFunc -code $code -funcRetType "void" -funcName "hogehoge" -funcParam "int x" -funcBody $funcBody

$code.ToString()