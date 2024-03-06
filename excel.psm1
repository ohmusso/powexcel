Write-Host $PSScriptRoot

Import-Module "${PSScriptRoot}\file.psm1"

$TableInfo = [PSCustomObject]@{
    PSTypeName = 'TableInfo.Object'
    StartRow = 0
    EndRow = 0
    StartColumn = 0
    EndColumn = 0
    PropertyRow = 0
    PropertyNames = @()
}

function New-TableInfo(){
    $TableInfo.psobject.Copy()
}

function Read-TableInfo{
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
    for( $column = $tableInfo.StartColumn; $column -lt $tableInfo.EndColumn; $column++ ){
        $propertyName = ""
        for( $row = $startHeaderRow; $row -lt $endHeaderRow; $row++ ){
            # ヘッダ行をループ
            $name = $rangeObj.cells($row, $column).text
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

function Read-Table{
    Param(
        $startCell = "A1",  # 表の見出しを含めた一番左上
        $rowOffset = 1,     # Currentregionでずれた分を補正
        $colOffset = 2,     # Currentregionでずれた分を補正
        $headerRow = 1,     # 表の見出し行数
        $sheet = $null,     # excel object
        $stringRange = ""
    )

    if( $null -eq $sheet ){
        Write-Error "no sheet"
        exit
    }

    $range
    if( $stringRange -eq "" ){
        $range = $sheet.Range($startCell).Currentregion
    }
    else{
        $range = $sheet.Range($stringRange)
    }

    [PSTypeName('TableInfo.Object')]$tableInfo = Read-TableInfo -rangeObj $range -rowOffset $rowOffset -colOffset $colOffset -headerRow $headerRow

    # 表をオブジェクト化
    $table = @()
    $rangeValue2 = $sheet.Range(
            $range.Cells($tableInfo.StartRow, $tableInfo.StartColumn),
            $range.Cells($tableInfo.EndRow, $tableInfo.EndColumn)
    ).Value2

    for( $row = 0; $row -lt ($tableInfo.EndRow - $tableInfo.StartRow); $row++){
        # オブジェクトを作成し、ヘッダ行をメンバとして追加
        $obj = New-Object -TypeName PSCustomObject
        foreach($propertyName in $tableInfo.PropertyNames){
            $obj | Add-Member -MemberType NoteProperty -Name $propertyName -Value "" # 全てのメンバは文字列で、空文字で初期化する。
        }

        # オブジェクトに読みだした行データを設定
        for( $column = 0; $column -lt ($tableInfo.EndColumn - $tableInfo.StartColumn); $column++){
            $obj.($tableInfo.PropertyNames[$column]) = $rangeValue2[($row + 1), ($column + 1)] # +1はValue2が1始まりのため
        }

        $table += $obj
    }

    # output
    $table
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

function Get-Excel{
    param(
        $startDir = "C:",
        $title = "excelファイルを選択してください",
        $isNew = $false
    )

    $excel = $null

    if( $isNew -eq $false ){
        $path = Open-FileDialog -startDir $startDir -title $title

        if( $null -eq $path ){
            Write-Error "ファイルが選択されませんでした"
            exit
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        try{
            [void]$excel.Workbooks.Open($path)
        }catch{
            $excel.Quit()
            $excel = $null
        }
    }
    else{
        if( $PSVersionTable.PSVersion.Revision -le 4046 ){
            # powershell 5.1 or less 
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Applicaiton")
        }
        else{
            # greater than powershell 5.1
            # definie Marshal2 in this script 
            $excel = [Marshal2]::GetActiveObject("Excel.Applicaiton")
        }
    }

    # return
    $excel
}
