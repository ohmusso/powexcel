using namespace System.Collections
using namespace System.Text

function Test-MatchValiables{
    param(
        $valiables,
        $str
    )

    foreach($v in $valiables){
        if( $str -match "${v}[^a-zA-Z1-9_]|${v}$" ){
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

$newLine = "`r`n"
#$tab = "`t"
function Add-CodeFunc{
    param(
        [StringBuilder]$code,
        $funcHeader = "",
        $funcRetType = "",
        $funcName = "",
        $funcParam = "",
        [StringBuilder]$funcBody
    )
	# $content = Get-Content -Path "${PSScriptRoot}\cedit_template\func_template.c" -Encoding Default -Raw
    [void]$code.Append($funcHeader)
    [void]$code.Append("${funcRetType} ${funcName}(${funcParam})${newLine}{${newLine}")
    [void]$code.Append("${funcBody}}")
    # $codeFunc.code = $codeFunc.code -replace "<funcReturnType>", $funcReturnType
    # $codeFunc.code = $codeFunc.code -replace "<funcName>", $funcName
    # $codeFunc.code = $codeFunc.code -replace "<funcParam>", $funcParam
#	$codeFunc.bodyStartIndex = [array]::IndexOf($codeFunc.code, "<funcBody>")
#	$codeFunc.bodyStartIndex = $codeFunc.bodyEndIndex
}

$tabLength = 4
function Add-CodeLine{
    param(
        [StringBuilder]$code,
        $line = "",
        $comment = "",
        $indent = 0,
        $tabNumBeginComment = 8
    )

	$lineLength = $line.Length
	if( $indent -gt 0 ){
		$line = ("`t" * $indent) + $line
		$lineLength += ($indent * $tabLength)
	}

	$commentBegin = $tabLength * $tabNumBeginComment
	if( $lineLength -lt $commentBegin ){
		$addTabNum = [math]::Floor(($commentBegin - $lineLength) / $tabLength) + 1
		$line = $line + ("`t" * $addTabNum) + "/* " + $comment + " */"
	}
	else{

	}

	[void]$code.Append($line)
	[void]$code.Append($newLine)
}
