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
