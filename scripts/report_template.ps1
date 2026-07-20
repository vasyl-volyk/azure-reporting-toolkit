param (
    [string]$Region,
    [string]$OutputPath
)

$object |  Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
