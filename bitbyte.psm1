# bits are aligned in big endian
function Get-BitsFromBytes{
    param(
        $bytes,
        $sigStartBit,
        $sigLength
    )
    if( ($sigLength -gt 32) -or ($sigLength -eq 0) ){
        Write-Error "invalid param"
        return
    }

    $sigStartByte = [math]::Floor(($sigStartBit / 8))
    $sigEndBit = $sigStartBit + $sigLength - 1
    $sigEndByte = [math]::Floor(($sigEndBit / 8))

    if( $sigEndByte -gt ($bytes.Count - 1) ){
        Write-Error "invalid param"
        return
    }

    $sigValue = 0
    if( $sigStartByte -ne $sigEndByte ){
        # バイト跨ぎあり
        # 後ろから処理
        $curBit = $sigEndBit
        $curByte = $sigEndByte

        # 開始Byteまで処理
        do{
            $numBit= $curBit - ($curByte * 8) + 1
            $curBit = $curBit - $numBit
            $shift = 8 - $numBit
            $mask = (1 -shl $numBit) - 1
            $extractedValue = ($bytes[$curByte] -shr $shift) -band $mask
            $sigValue = $sigValue -bor ($extractedValue  -shl ($curBit - $sigStartBit + 1))

            $curByte--

        } while($curByte -gt $sigStartByte)

        # 開始Byte処理
        $numBit = $curBit - $sigStartBit + 1
        $curBit = $curBit - $numBit
        $shift = 0
        $mask = (1 -shl $numBit) - 1
        $extractedValue = ($bytes[$curByte] -shr $shift) -band $mask
        $sigValue = $sigValue -bor ($extractedValue  -shl ($curBit - $sigStartBit + 1))
    }
    else{
        # バイト跨ぎなし
        $shift = 7 - $sigEndBit
        $mask = (1 -shl $sigLength) - 1
        $sigValue = ($bytes[$sigStartByte] -shr $shift) -band $mask
    }

    $sigValue
}

# bits are aligned in little endian
function Get-BitsFromBytesLittle{
    param(
        $bytes,
        $sigStartBit,
        $sigLength
    )
    if( ($sigLength -gt 32) -or ($sigLength -eq 0) ){
        Write-Error "invalid param"
        return
    }

    $sigStartByte = [math]::Floor(($sigStartBit / 8))
    $sigEndBit = $sigStartBit + $sigLength - 1
    $sigEndByte = [math]::Floor(($sigEndBit / 8))

    if( $sigEndByte -gt ($bytes.Count - 1) ){
        Write-Error "invalid param"
        return
    }

    $sigValue = 0
    if( $sigStartByte -ne $sigEndByte ){
        # バイト跨ぎあり
        # 後ろから処理
        $curBit = $sigEndBit
        $curByte = $sigEndByte

        # 開始Byteまで処理
        do{
            $numBit= $curBit - ($curByte * 8) + 1
            $curBit = $curBit - $numBit
            $shift = 0
            $mask = (1 -shl $numBit) - 1
            $extractedValue = ($bytes[$curByte] -shr $shift) -band $mask
            $sigValue = $sigValue -bor ($extractedValue  -shl ($curBit - $sigStartBit + 1))

            $curByte--

        } while($curByte -gt $sigStartByte)

        # 開始Byte処理
        $numBit = $curBit - $sigStartBit + 1
        $curBit = $curBit - $numBit
        $shift = 8 - $numBit
        $mask = (1 -shl $numBit) - 1
        $extractedValue = ($bytes[$curByte] -shr $shift) -band $mask
        $sigValue = $sigValue -bor ($extractedValue  -shl ($curBit - $sigStartBit + 1))
    }
    else{
        # バイト跨ぎなし
        $shift = $sigStartBit
        $mask = (1 -shl $sigLength) - 1
        $sigValue = ($bytes[$sigStartByte] -shr $shift) -band $mask
    }

    $sigValue
}
