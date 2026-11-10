param(
  [Parameter(Mandatory = $true)]
  [string]$Destination
)

$ErrorActionPreference = 'Stop'

try {
  New-Item -ItemType Directory -Force -Path $Destination | Out-Null

  Write-Host "Recherche de la derniere LTS Node.js..."
  $releases = Invoke-RestMethod -Uri 'https://nodejs.org/dist/index.json'
  $release = $releases | Where-Object { $_.lts } | Select-Object -First 1
  if (-not $release) { throw 'Aucune version LTS trouvee sur nodejs.org.' }

  $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
  $version = $release.version
  $folder = "node-$version-win-$arch"
  $zipName = "$folder.zip"
  $url = "https://nodejs.org/dist/$version/$zipName"
  $zipPath = Join-Path $Destination $zipName

  if (Test-Path (Join-Path $Destination "$folder\node.exe")) {
    Write-Host "Deja present : $folder"
    exit 0
  }

  Write-Host "Telechargement $url"
  Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing

  Write-Host "Extraction vers $Destination"
  Expand-Archive -LiteralPath $zipPath -DestinationPath $Destination -Force
  Remove-Item -LiteralPath $zipPath -Force

  $nodeExe = Join-Path $Destination "$folder\node.exe"
  if (-not (Test-Path $nodeExe)) { throw "node.exe introuvable apres extraction ($folder)." }

  Write-Host "Node.js $version ($($release.lts)) installe sans droits admin."
  exit 0
} catch {
  Write-Host $_.Exception.Message
  exit 1
}
