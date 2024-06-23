Import-Module "${PSScriptRoot}/../../xmledit.psm1" -Force

$xml = [xml](Get-Content ./sample.arxml)

$xmlNew = [xml](Get-Content ./addnode.arxml)
$newNode= $xmlNew.AUTOSAR.'AR-PACKAGE'
$newNode.'SHORT-NAME' = "Motor"

$node = Get-XmlNode -xml $xml -nodeName "AR-PACKAGE" -key "SHORT-NAME" -value "Demo" -namespace "http://autosar.org/schema/r4.0"

Add-XmlNodeFromAnotherDoc -xml $xml -parentNode $node -addNode $newNode

$xml.Save("sample_add.arxml")
