function Open-FileDialog{
    param(
        $startDir = "C:",
        $title = "ファイルを選択してください"    
    )
    
    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    if( $null -eq $dialog ){
        Write-Error "ファイルダイアログの起動に失敗しました"
        exit
    }

    $dialog.Title = $title
    $dialog.Filter = "全てのファイル(*.*)|*.*"
    $dialog.InitialDirectory = $startDir

    if( $dialog.ShowDialog() -eq "OK" ){
        # output
        $dialog.FileName 
    }
}

function Open-FolderDialog{
    param(
        $startDir = "C:",
        $title = "フォルダを選択してください"    
    )
    
    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog

    if( $null -eq $dialog ){
        Write-Error "フォルダダイアログの起動に失敗しました"
        exit
    }

    $dialog.Description = $title
    $dialog.SelectedPath = $startDir

    if( $dialog.ShowDialog() -eq "OK" ){
        # output
        $dialog.SelectedPath
    }
}
