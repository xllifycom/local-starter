$uuid = [System.Guid]::NewGuid().ToString()
$path = Join-Path $PSScriptRoot "xllify.json"
$json = Get-Content $path -Raw | ConvertFrom-Json
$json.app_id = $uuid
$json | ConvertTo-Json -Depth 10 | Set-Content $path
Write-Host "app_id set to $uuid"
