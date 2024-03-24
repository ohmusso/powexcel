using namespace System.Collections

Import-Module "${PSScriptRoot}\../../cedit.psm1" -Force

$files = @(
    "test1.txt",
    "test2.txt",
    "test3.txt"
)

$valiables = @'
abc_edf
hugahuga
'@

# ”z—ñ‚É•ÏŠ·
$valiables = $valiables -split "`r`n"

Remove-Line $files $valiables
