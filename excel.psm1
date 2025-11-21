Write-Host $PSScriptRoot

Import-Module "${PSScriptRoot}\file.psm1"

$TableInfo = [PSCustomObject]@{
	PSTypeName    = 'TableInfo.Object'
	StartRow      = 0
	EndRow        = 0
	StartColumn   = 0
	EndColumn     = 0
	PropertyRow   = 0
	PropertyNames = @()
}

$ArrayInfo = [PSCustomObject]@{
	PSTypeName = 'ArrayInfo.Object'
	RowCount   = 0
	ColCount   = 0
	Array      = $null
}

function New-TableInfo() {
	$TableInfo.psobject.Copy()
}

function Read-TableInfo {
	Param(
		$rangeObj,
		$rowOffset = 2,
		$colOffset = 2,
		$headerRow = 1,
		$headerDelim = ":"
	)
    
	$tableInfo = $TableInfo.psobject.copy()

	$tableInfo.StartRow = $rowOffset + $headerRow + 1  # データ行の開始 +1 はRange.cells()が1始まりの為
	$tableInfo.EndRow = $rangeObj.Rows.Count + 1       # データ行の終了 +1 はRange.cells()が1始まりの為
	$tableInfo.StartColumn = $colOffset + 1            # +1 はRange.cells()が1始まりの為
	$tableInfo.EndColumn = $rangeObj.Columns.Count + 1 # +1 はRange.cells()が1始まりの為

	# ヘッダの行数
	$startHeaderRow = $rowOffset + 1
	$endHeaderRow = $startHeaderRow + $headerRow

	$StackPropertyName = New-Object String[] $headerRow # ヘッダ行が複数行の場合、各行の文字列を連結して一つの列名とする
	for ( $column = $tableInfo.StartColumn; $column -lt $tableInfo.EndColumn; $column++ ) {
		$propertyName = ""
		for ( $row = $startHeaderRow; $row -lt $endHeaderRow; $row++ ) {
			# ヘッダ行をループ
			$name = $rangeObj.cells($row, $column).text
			if ( ($rangeObj.cells($row, $column).MergeCells -eq $true) -and ($name -ne "") ) {
				# 結合セルかつ空でないの場合
				# 結合セルの左端。ヘッダの親要素とする。
				$StackPropertyName[$row - $startHeaderRow] = $name + $headerDelim 
			}
			else {
				# ヘッダの子要素。
				$StackPropertyName[$row - $startHeaderRow] = ""
				$propertyName = $name
			}
		}

		if ( $propertyName -eq "" ) {
			# 空白列の場合
			$propertyName = "reserved_" + $column 
		}
		else {
			# スタックの文字を連結して一つの列名とする
			$propertyName = [string]::Join("", $StackPropertyName) + $propertyName 
		}
		$tableInfo.PropertyNames += $propertyName 
	}

	$tableInfo
}

function Read-Table {
	Param(
		$startCell = "A1",  # 表の見出しを含めた一番左上
		$rowOffset = 1,     # Currentregionでずれた分を補正
		$colOffset = 2,     # Currentregionでずれた分を補正
		$headerRow = 1,     # 表の見出し行数
		$sheet = $null,     # excel object
		$stringRange = ""
	)

	if ( $null -eq $sheet ) {
		Write-Error "no sheet"
		exit
	}

	if ( $stringRange -eq "" ) {
		$range = $sheet.Range($startCell).Currentregion
	}
	else {
		$range = $sheet.Range($stringRange)
	}

	[PSTypeName('TableInfo.Object')]$tableInfo = Read-TableInfo -rangeObj $range -rowOffset $rowOffset -colOffset $colOffset -headerRow $headerRow

	# 表をオブジェクト化
	$table = @()
	$rangeValue2 = $sheet.Range(
		$range.Cells($tableInfo.StartRow, $tableInfo.StartColumn),
		$range.Cells($tableInfo.EndRow, $tableInfo.EndColumn)
	).Value2

	# テンプレートオブジェクトを作成。ヘッダ行をメンバとして追加
	$tableObj = New-Object -TypeName PSCustomObject
	foreach ($propertyName in $tableInfo.PropertyNames) {
		$tableObj | Add-Member -MemberType NoteProperty -Name $propertyName -Value "" # 全てのメンバは文字列で、空文字で初期化する。
	}

	# テンプレートからオブジェクトを作成して読み出したデータを設定する
	for ( $row = 0; $row -lt ($tableInfo.EndRow - $tableInfo.StartRow); $row++) {

		$obj = $tableObj.psobject.Copy()

		# オブジェクトに読みだした行データを設定
		for ( $column = 0; $column -lt ($tableInfo.EndColumn - $tableInfo.StartColumn); $column++) {
			$obj.($tableInfo.PropertyNames[$column]) = $rangeValue2[($row + 1), ($column + 1)] # +1はValue2が1始まりのため
		}

		$table += $obj
	}

	# output
	$table
}

function Copy-Table {
	Param(
		$fromSheet = $null,
		$fromRangeStr = "A1:B1",
		$destSheet = $null,
		$destCellStr = "D2",
		$filterCol = $null,
		$filterVal = $null
	)

	if ( ($null -eq $fromSheet) -or ($null -eq $destSheet) ) {
		Write-Error "no sheet object"
		exit
	}

	# コピー元のテーブル範囲を調べる
	$cellStrs = $fromRangeStr -split ":"

	$fromStartRow = $fromSheet.Range($cellStrs[0]).Row
	$fromStartCol = $fromSheet.Range($cellStrs[0]).Column

	$fromEndCol = $fromSheet.Range($cellStrs[1]).Column
	$fromEndRow = $fromStartRow
	while ($null -ne $fromSheet.Cells($fromEndRow, $fromEndCol).Value()) {
		$fromEndRow++
	}

	# コピー先のセル情報
	$destStartRow = $destSheet.Range($destCellStr).Row
	$destStartCol = $destSheet.Range($destCellStr).Column
	$destEndRow = $destStartRow + ($fromEndRow - $fromStartRow)
	$destEndCol = $destStartCol + ($fromEndCol - $fromStartCol)

	# コピー元の数式を値で上書き
	$fromSheet.Activate()
    $v = $fromSheet.Range($fromSheet.Cells($fromStartRow, $fromStartCol), $fromSheet.Cells($fromEndRow, $fromEndCol)).Value()
	$fromSheet.Range($fromSheet.Cells($fromStartRow, $fromStartCol), $fromSheet.Cells($fromEndRow, $fromEndCol)).Value() = $v

    # フィルタ処理
	$fromRagne = $fromSheet.Range($fromSheet.Cells($fromStartRow, $fromStartCol), $fromSheet.Cells($fromEndRow-1, $fromEndCol))
    $fromRagne.EntireRow.Hidden = $false
    $fromRagne.EntireColumn.Hidden = $false
	if ( $null -ne $filterCol ) {
		$fromRagne.AutoFilter($filterCol, $filterVal)
		$filterdRange = $fromRagne.SpecialCells(12)
        $areaNum = $filterdRange.Areas.Count
        $destSheet.Activate()
        $destRowOffset = $destStartRow
        for( $i = 0; $i -lt $areaNum; $i++ ){
    	    # Area毎に値と書式をコピー Value()の引数 11: xlRangeValueXMLSpreadsheet
            $valueXml = $filterdRange.Areas.Item($i + 1).Value(11) # Item begin from 1
            $rowCount = $filterdRange.Areas.Item($i + 1).Rows.Count
	        $destSheet.Range($destSheet.Cells($destRowOffset, $destStartCol), $destSheet.Cells($destRowOffset + $rowCount - 1, $destEndCol)).Value(11) = $valueXml
            $destRowOffset += $rowCount
        }
	}else{
	    # 値と書式をコピー Value()の引数 11: xlRangeValueXMLSpreadsheet
	    $destSheet.Activate()
        $valueXml = $fromRagne.Value(11)
	    $destSheet.Range($destSheet.Cells($destStartRow, $destStartCol), $destSheet.Cells($destEndRow - 1, $destEndCol)).Value(11) = $valueXml
    }

}

Add-Type -TypeDefinition @'
using System;
using System.Runtime;
using System.Runtime.InteropServices;
public static class Marshal2
{
    internal const String OLEAUT32 = "oleaut32.dll";
    internal const String OLE32 = "ole32.dll";

    public static Object GetActiveObject(String progID)
    {
        Object obj = null;
        Guid clsid;

        // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
        // CLSIDFromProgIDEx doesn't exist.
        try
        {
            CLSIDFromProgIDEx(progID, out clsid);
        }
        //            catch
        catch (Exception)
        {
            CLSIDFromProgID(progID, out clsid);
        }

        GetActiveObject(ref clsid, IntPtr.Zero, out obj);
        return obj;
    }

    //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
    [DllImport(OLE32, PreserveSig = false)]
    private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

    //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
    [DllImport(OLE32, PreserveSig = false)]
    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

    //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
    [DllImport(OLEAUT32, PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);

}
'@

function Get-Excel {
	param(
		$startDir = "C:",
		$title = "excelファイルを選択してください",
		$isOpen = $true
	)

	$excel = $null

	if ( $isOpen -eq $true ) {
		$path = Open-FileDialog -startDir $startDir -title $title

		if ( $null -eq $path ) {
			Write-Error "ファイルが選択されませんでした"
			exit
		}

		$excel = New-Object -ComObject Excel.Application
		$excel.Visible = $false
		try {
			[void]$excel.Workbooks.Open($path)
		}
		catch {
			$excel.Quit()
			$excel = $null
		}
	}
	else {
		if (($PSVersionTable.PSVersion.Major -le 5) -and ($PSVersionTable.PSVersion.Minor -le 1) ) {
			# powershell 5.1 or less 
			$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
		}
		else {
			# greater than powershell 5.1
			# definie Marshal2 in this script 
			$excel = [Marshal2]::GetActiveObject("Excel.Application")
		}
	}

	# return
	$excel
}

function Convert-ObjsToArray {
	param(
		$objs = $null,
		$isHeader = $true
	)

	$arrayInfo = $ArrayInfo.psobject.copy()

	if ( $isHeader -eq $true ) {
		$arrayInfo.RowCount = $objs.Count + 1
	}
	else {
		$arrayInfo.RowCount = $objs.Count
	}
	$arrayInfo.ColCount = ($objs | Get-Member -MemberType NoteProperty).Count

	$properties = $objs[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

	$array = New-Object 'Object[,]' $arrayInfo.RowCount, $arrayInfo.ColCount
    
	$rowOffset = 0
	if ( $isHeader -eq $true ) {
		for ( $col = 0; $col -lt $arrayInfo.ColCount; $col++ ) {
			$array[0, $col] = [String]$properties[$col]
		}

		$rowOffset = 1
	}

	for ( $row = 0; $row -lt $objs.Count; $row++ ) {
		for ( $col = 0; $col -lt $arrayInfo.ColCount; $col++ ) {
			$array[($row + $rowOffset), $col] = [String]$objs[$row].($properties[$col])
		}
	}
	$arrayInfo.Array = $array

	$arrayInfo
}

function Write-Table {
	param (
		$startCell = "A1",
		$sheet = $null,
		$objs = $null,
		$isHeader = $true
	)

	if ( $null -eq $sheet ) {
		Write-Error "sheet is null"
		exit
	}

	if ( $null -eq $objs ) {
		Write-Error "Objs is null"
	}

	[PSTypeName('ArrayInfo.Object')]$arrayInfo = Convert-ObjsToArray -objs $objs -isHeader $isHeader

	# 書き込み開始位置
	$startRow = $sheet.Range($startCell).Row
	$startCol = $sheet.Range($startCell).Column
	# 書き込み終了位置
	$endRow = $startRow + $arrayInfo.RowCount - 1 # 0始まり
	$endCol = $startCol + $arrayInfo.ColCount - 1 # 0始まり
	# 書き込み
	$sheet.Range($sheet.Cells($startRow, $startCol), $sheet.Cells($endRow, $endCol)).Value2 = $arrayInfo.Array
}

function New-Excel {
	param(
		$filePath = "book1.xlsx",
		$startCell,
		$objects
	)


	$excel = New-Object -ComObject Excel.Application
	$excel.DisplayAlerts = $FALSE

	$book = $excel.Workbooks.Add();

	$book.SaveAs($filePath);

	$excel.Quit()
	$book = $null
	$excel = $null
}
