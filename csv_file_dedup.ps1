# Read the dataset and dedup the contents
# 3 arguments should be passed on the command line
#  input filename
#  output filename
#  sort field (note: a single sort field may be specified)

if ($args.Count -ne 3) {
    Write-Host("Expecting 3 parameters: input, output filenames and sort field")
    exit 1
}
#
# Map the network path to a drive letter
$InPath = Split-Path -Parent -Path $args[0]
$InFile = Split-Path -Leaf -Path $args[0]
$OutPath = Split-Path -Parent -Path $args[1]
$OutFile = Split-Path -Leaf -Path $args[1]

New-PSDrive -Name I -PSProvider "FileSystem" -Root $InPath
New-PSDrive -Name O -PSProvider "FileSystem" -Root $OutPath

Write-Host("Importing $($args[0])")
Write-Host("Exporting unique rows to $($args[1])")
Write-Host("Sorting on $($args[2])")

$P = Import-Csv -Path I:$InFile
#Write-Host("Reading $($P.Count) rows")
$P | Sort-Object -Descending -Unique -Property $args[2] | Export-Csv -NoTypeInformation -Path O:$OutFile
