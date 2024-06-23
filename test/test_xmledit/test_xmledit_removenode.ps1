Import-Module "${PSScriptRoot}/../../xmledit.psm1" -Force

$xml = [xml](Get-Content ./sample.arxml)

Remove-XmlNode -xml $xml -nodeName "AR-PACKAGE" -key "SHORT-NAME" -value "Door" -namespace "http://autosar.org/schema/r4.0"

$xml.Save("sample_noderemove.arxml")
