Import-Module "${PSScriptRoot}/../bitbyte.psm1" -Force

# Get-BitsFromBytes
$datas = @(
    255, 255, 255, 255, 255, 255, 255, 255
)
-1 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 0 -sigLength 32)
-1 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 32 -sigLength 32)

-1 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 31 -sigLength 32)

$datas = @(
    0, 0, 0, 255, 255, 255, 255, 255
)
255 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 0 -sigLength 32)

$datas = @(
    0, 0, 255, 255, 255, 255, 255, 255
)
65535 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 0 -sigLength 32)

$datas = @(
    0, 255, 255, 255, 255, 255, 255, 255
)
16777215 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 0 -sigLength 32)

$datas = @(
    0, 0, 0, 1, 128, 0, 0, 0
)
3 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 31 -sigLength 2)
1 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 31 -sigLength 1)
1 -eq (Get-BitsFromBytes -bytes $datas -sigStartBit 32 -sigLength 1)

# Get-BitsFromBytesLittle
$datas = @(
    255, 3, 255, 255, 255, 255, 255, 255
)
1023 -eq (Get-BitsFromBytesLittle -bytes $datas -sigStartBit 0 -sigLength 10)

$datas = @(
    0, 0, 24, 0, 0, 0, 0, 0
)
1 -eq (Get-BitsFromBytesLittle -bytes $datas -sigStartBit 19 -sigLength 1)
1 -eq (Get-BitsFromBytesLittle -bytes $datas -sigStartBit 20 -sigLength 1)

