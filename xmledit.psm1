<#
.SYNOPSIS
#remove xml node

.DESCRIPTION
remove xml node

.PARAMETER valiables
Parameter description

.PARAMETER str
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Remove-XmlNode{
    param(
        [xml]$xml,
        [String]$nodeName,
        [String]$key,
        $value,
        [String]$namespace
    )
	$ns = @{ns = $namespace}
	$xPath = "//ns:$nodeName[ns:$key='$value']"
	$selectNodeInfo = Select-Xml -Xml $xml -XPath $xPath -Namespace $ns
    $node = $selectNodeInfo.Node
	$node.ParentNode.RemoveChild($node)
}

<#
.SYNOPSIS
#get xml node

.DESCRIPTION
get xml node

.PARAMETER valiables
Parameter description

.PARAMETER str
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-XmlNode{
    param(
        [xml]$xml,
        [String]$nodeName,
        [String]$key,
        $value,
        [String]$namespace
    )
	$ns = @{ns = $namespace}
	$xPath = "//ns:$nodeName[ns:$key='$value']"
	$selectNodeInfo = Select-Xml -Xml $xml -XPath $xPath -Namespace $ns
    return $selectNodeInfo.Node
}

<#
.SYNOPSIS
#Add xml node

.DESCRIPTION
Add xml node

.PARAMETER valiables
Parameter description

.PARAMETER str
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function Add-XmlNodeFromAnotherDoc{
    param(
        [xml]$xml,   # same document
        $parentNode, # same document
        $addNode     # another document
    )
    # import node from another document
    $importedNode = $xml.ImportNode($addNode, $true)
    # imported node can be appended
    $parentNode.AppendChild($importedNode)
}
