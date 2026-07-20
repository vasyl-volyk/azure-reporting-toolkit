param (
    [string]$FilePath,
    [string]$Container,
    [string]$ConnectionString,
    [string]$BlobPath
)

az storage blob upload `
  --file $FilePath `
  --container-name $Container `
  --name $BlobPath `
  --overwrite true `
  --connection-string $ConnectionString
