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

$ArrayInfo = [PSCustomObject]@{
    PSTypeName = 'ArrayInfo.Object'
    RowCount = 0
    ColCount = 0
    Array = $null
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

    $tableInfo.StartRow = $rowOffset + $headerRow + 1  # �f�[�^�s�̊J�n +1 ��Range.cells()��1�n�܂�̈�
    $tableInfo.EndRow = $rangeObj.Rows.Count + 1       # �f�[�^�s�̏I�� +1 ��Range.cells()��1�n�܂�̈�
    $tableInfo.StartColumn = $colOffset + 1            # +1 ��Range.cells()��1�n�܂�̈�
    $tableInfo.EndColumn = $rangeObj.Columns.Count + 1 # +1 ��Range.cells()��1�n�܂�̈�

    # �w�b�_�̍s��
    $startHeaderRow = $rowOffset + 1
    $endHeaderRow = $startHeaderRow + $headerRow

    $StackPropertyName = New-Object String[] $headerRow # �w�b�_�s�������s�̏ꍇ�A�e�s�̕������A�����Ĉ�̗񖼂Ƃ���
    for( $column = $tableInfo.StartColumn; $column -lt $tableInfo.EndColumn; $column++ ){
        $propertyName = ""
        for( $row = $startHeaderRow; $row -lt $endHeaderRow; $row++ ){
            # �w�b�_�s�����[�v
            $name = $rangeObj.cells($row, $column).text
            if( ($rangeObj.cells($row, $column).MergeCells -eq $true) -and ($name -ne "") ){ # �����Z������łȂ��̏ꍇ
                # �����Z���̍��[�B�w�b�_�̐e�v�f�Ƃ���B
                $StackPropertyName[$row - $startHeaderRow] = $name + $headerDelim 
            }
            else{
                # �w�b�_�̎q�v�f�B
                $StackPropertyName[$row - $startHeaderRow] = ""
                $propertyName = $name
            }
        }

        if( $propertyName -eq "" ){
            # �󔒗�̏ꍇ
            $propertyName = "reserved_" + $column 
        }
        else{
            # �X�^�b�N�̕�����A�����Ĉ�̗񖼂Ƃ���
            $propertyName = [string]::Join("", $StackPropertyName) + $propertyName 
        }
        $tableInfo.PropertyNames += $propertyName 
    }

    $tableInfo
}

function Read-Table{
    Param(
        $startCell = "A1",  # �\�̌��o�����܂߂���ԍ���
        $rowOffset = 1,     # Currentregion�ł��ꂽ����␳
        $colOffset = 2,     # Currentregion�ł��ꂽ����␳
        $headerRow = 1,     # �\�̌��o���s��
        $sheet = $null,     # excel object
        $stringRange = ""
    )

    if( $null -eq $sheet ){
        Write-Error "no sheet"
        exit
    }

    if( $stringRange -eq "" ){
        $range = $sheet.Range($startCell).Currentregion
    }
    else{
        $range = $sheet.Range($stringRange)
    }

    [PSTypeName('TableInfo.Object')]$tableInfo = Read-TableInfo -rangeObj $range -rowOffset $rowOffset -colOffset $colOffset -headerRow $headerRow

    # �\���I�u�W�F�N�g��
    $table = @()
    $rangeValue2 = $sheet.Range(
            $range.Cells($tableInfo.StartRow, $tableInfo.StartColumn),
            $range.Cells($tableInfo.EndRow, $tableInfo.EndColumn)
    ).Value2

    # �e���v���[�g�I�u�W�F�N�g���쐬�B�w�b�_�s�������o�Ƃ��Ēǉ�
    $tableObj = New-Object -TypeName PSCustomObject
    foreach($propertyName in $tableInfo.PropertyNames){
        $tableObj | Add-Member -MemberType NoteProperty -Name $propertyName -Value "" # �S�Ẵ����o�͕�����ŁA�󕶎��ŏ���������B
    }

    # �e���v���[�g����I�u�W�F�N�g���쐬���ēǂݏo�����f�[�^��ݒ肷��
    for( $row = 0; $row -lt ($tableInfo.EndRow - $tableInfo.StartRow); $row++){

        $obj = $tableObj.psobject.Copy()

        # �I�u�W�F�N�g�ɓǂ݂������s�f�[�^��ݒ�
        for( $column = 0; $column -lt ($tableInfo.EndColumn - $tableInfo.StartColumn); $column++){
            $obj.($tableInfo.PropertyNames[$column]) = $rangeValue2[($row + 1), ($column + 1)] # +1��Value2��1�n�܂�̂���
        }

        $table += $obj
    }

    # output
    $table
}

function Copy-Table{
    Param(
        $fromSheet = $null,
        $fromRangeStr = "A1:B1",
        $destSheet = $null,
        $destCellStr = "D2"
    )

    if( ($null -eq $fromSheet) -or ($null -eq $destSheet) ){
        Write-Error "no sheet object"
        exit
    }

    # �R�s�[���̃e�[�u���͈͂𒲂ׂ�
    $cellStrs = $fromRangeStr -split ":"

    $fromStartRow = $fromSheet.Range($cellStrs[0]).Row
    $fromStartCol = $fromSheet.Range($cellStrs[0]).Column

    $fromEndCol = $fromSheet.Range($cellStrs[1]).Column
    $fromEndRow = $fromStartRow
    while($fromSheet.Cells($fromEndRow, $fromEndCol).Value() -ne $null){
        $fromEndRow++
    }

    # �R�s�[��̃Z�����
    $destStartRow = $destSheet.Range($destCellStr).Row
    $destStartCol = $destSheet.Range($destCellStr).Column
    $destEndRow = $destStartRow + ($fromEndRow - $fromStartRow)
    $destEndCol = $destStartCol + ($fromEndCol - $fromStartCol)

    # �R�s�[���̐�����l�ŏ㏑��
    $fromSheet.Activate()
    $fromSheet.Range($fromSheet.Cells($fromStartRow, $fromStartCol), $fromSheet.Cells($fromEndRow, $fromEndCol)).Value() = $fromSheet.Range($fromSheet.Cells($fromStartRow, $fromStartCol), $fromSheet.Cells($fromEndRow, $fromEndCol)).Value()

    # �l�Ə������R�s�[ Value()�̈��� 11: xlRangeValueXMLSpreadsheet
    $destSheet.Activate()
    $destSheet.Range($destSheet.Cells($destStartRow, $destStartCol), $destSheet.Cells($destEndRow, $destEndCol)).Value(11) = $fromSheet.Range($fromSheet.Cells($fromStartRow, $fromStartCol), $fromSheet.Cells($fromEndRow, $fromEndCol)).Value(11)
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
        $title = "excel�t�@�C����I�����Ă�������",
        $isOpen = $true
    )

    $excel = $null

    if( $isOpen -eq $true ){
        $path = Open-FileDialog -startDir $startDir -title $title

        if( $null -eq $path ){
            Write-Error "�t�@�C�����I������܂���ł���"
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
        if(($PSVersionTable.PSVersion.Major -le 5) -and ($PSVersionTable.PSVersion.Minor -le 1) ){
            # powershell 5.1 or less 
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        }
        else{
            # greater than powershell 5.1
            # definie Marshal2 in this script 
            $excel = [Marshal2]::GetActiveObject("Excel.Application")
        }
    }

    # return
    $excel
}

function Convert-ObjsToArray{
    param(
        $objs = $null,
        $isHeader = $true
    )

    $arrayInfo= $ArrayInfo.psobject.copy()

    if( $isHeader -eq $true ){
        $arrayInfo.RowCount = $objs.Count + 1
    }
    else{
        $arrayInfo.RowCount = $objs.Count
    }
    $arrayInfo.ColCount = ($objs | Get-Member -MemberType NoteProperty).Count

    $properties = $objs[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

    $array = New-Object 'Object[,]' $arrayInfo.RowCount, $arrayInfo.ColCount
    
    $rowOffset = 0
    if( $isHeader -eq $true ){
        for( $col = 0; $col -lt $arrayInfo.ColCount; $col++ ){
            $array[0, $col] = [String]$properties[$col]
        }

        $rowOffset = 1
    }

    for( $row = 0; $row -lt $objs.Count; $row++ ){
        for( $col = 0; $col -lt $arrayInfo.ColCount; $col++ ){
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

    if( $null -eq $sheet ){
        Write-Error "sheet is null"
        exit
    }

    if( $null -eq $objs ){
        Write-Error "Objs is null"
    }

    [PSTypeName('ArrayInfo.Object')]$arrayInfo = Convert-ObjsToArray -objs $objs -isHeader $isHeader

    # �������݊J�n�ʒu
    $startRow = $sheet.Range($startCell).Row
    $startCol = $sheet.Range($startCell).Column
    # �������ݏI���ʒu
    $endRow =  $startRow + $arrayInfo.RowCount - 1 # 0�n�܂�
    $endCol =  $startCol + $arrayInfo.ColCount - 1 # 0�n�܂�
    # ��������
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
