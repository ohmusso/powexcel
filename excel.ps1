. Join-Path $PSScriptRoot "file.ps1"

function Read-TableInfo{
    Param(
        $rangeObj,
        $rowOffset = 2,
        $colOffset = 2,
        $headerRow = 1,
        $headerDelim = ":"
    )
    
    $tableInfo = [PSCustomObject]@{
        StartRow = 0
        EndRow = 0
        StartColumn = 0
        EndColumn = 0
        PropertyRow = 0
        PropertyNames = @()
    }

    $tableInfo.StartRow = $rowOffset + $headerRow + 1 # データ行の開始 +1 はRange.cells()が1始まりの為
    $tableInfo.EndRow = $rangeObj.Row.Count + 1       # データ行の終了 +1 はRange.cells()が1始まりの為
    $tableInfo.StartColumn = $colOffset + 1           # +1 はRange.cells()が1始まりの為
    $tableInfo.EndColumn = $rangeObj.Column.Count + 1 # +1 はRange.cells()が1始まりの為

    # ヘッダの行数
    $startHeaderRow = $rowOffset + 1
    $endHeaderRow = $startHeaderRow + $headerRow

    $StackPropertyName = New-Object String[] $headerRow # ヘッダ行が複数行の場合、各行の文字列を連結して一つの列名とする
    for( $column = $tableInfo.StartColumn; $column -lt $tableInfo.EndColumn; $column++ ){
        $propertyName = ""
        for( $row = $startHeaderRow; $row -lt $endHeaderRow; $row++ ){
            # ヘッダ行をループ
            $name = $rangeObj.cells($row, $column).test
            if( ($rangeObj.cells($row, $column).MergeCells -eq $true) -and ($name -ne "") ){ # 結合セルかつ空でないの場合
                # 結合セルの左端。ヘッダの親要素とする。
                $StackPropertyName[$row - $startHeaderRow] = $name + $headerDelim 
            }
            else{
                # ヘッダの子要素。
                $StackPropertyName[$row - $startHeaderRow] = ""
                $propertyName = $name
            }
        }

        if( $propertyName -eq "" ){
            # 空白列の場合
            $propertyName = "reserved_" + $column 
        }
        else{
            # スタックの文字を連結して一つの列名とする
            $propertyName = [string]::Join("", $StackPropertyName) + $propertyName 
        }
        $tableInfo.PropertyNames += $propertyName 
    }

    $tableInfo
}
