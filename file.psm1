function Open-FileDialog{
    param(
        $startDir = "C:",
        $title = "�t�@�C����I�����Ă�������"    
    )
    
    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    if( $null -eq $dialog ){
        Write-Error "�t�@�C���_�C�A���O�̋N���Ɏ��s���܂���"
        exit
    }

    $dialog.Title = $title
    $dialog.Filter = "�S�Ẵt�@�C��(*.*)|*.*"
    $dialog.InitialDirectory = $startDir

    if( $dialog.ShowDialog() -eq "OK" ){
        # output
        $dialog.FileName 
    }
}

function Open-FolderDialog{
    param(
        $startDir = "C:",
        $title = "�t�H���_��I�����Ă�������"    
    )
    
    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog

    if( $null -eq $dialog ){
        Write-Error "�t�H���_�_�C�A���O�̋N���Ɏ��s���܂���"
        exit
    }

    $dialog.Description = $title
    $dialog.SelectedPath = $startDir

    if( $dialog.ShowDialog() -eq "OK" ){
        # output
        $dialog.SelectedPath
    }
}
