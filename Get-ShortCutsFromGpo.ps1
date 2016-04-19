param
(
	[Parameter(Mandatory = $true,
			   Position = 0)]
	[string]$outFile
)
Import-Module GroupPolicy


[xml]$data = Get-GPOReport -All -ReportType xml
$holder = foreach ($gpo in $data.gpos.gpo)
{
	$Gponame = $gpo.name
	foreach ($gpoShortcut in ($gpo.user.ExtensionData | Where-Object{ $_.name -like "Shortcuts" }).extension.shortcutsettings.shortcut)
		{
		$properties = [Ordered]@{
			"ApplicationName" = $gpoShortcut.name;
			"GPOName" = $Gponame;
			"ShortcutPath" = $gpoShortcut.properties.shortcutpath;
			"TargetType" = $gpoShortcut.properties.targettype;
			"StartIn" = $gpoShortcut.properties.startin;
			"TargetPath" = $gpoShortcut.properties.targetpath
		}
		$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $properties
		$obj
	}
}
$holder | Export-Csv $outFile -NoTypeInformation

