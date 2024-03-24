using namespace System.Collections

function Test-MatchValiables{
    param(
        $valiables,
        $str
    )

    foreach($v in $valiables){
        if( $str -match "${v}[^a-zA-Z1-9_]" ){
            return $true
        }
    }

    return $false
}

function Remove-Line {
    param(
        $files = @(),
        $valiables = @()
    )

    foreach( $file in $files ){
        $text = [ArrayList](Get-Content -Path $file -Encoding Default)

        if( $null -eq $text ){
            Write-Error "no file"
            exit
        }

        $i = 0
        do{
            $ret = Test-MatchValiables -valiables $valiables -str $text[$i]
            if( $ret -eq $true ){
                Write-Host ("Remove, ${file}: ${i}L: " + $text[$i])
                $text.RemoveAt($i)
            }

            $i++
        } while($null -ne $text[$i])

        #$text =[String]$text 
        Set-Content -Path $file -Value $text -Force
    }
}
