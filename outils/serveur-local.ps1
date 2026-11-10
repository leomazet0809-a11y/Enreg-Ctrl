# Serveur web local pour "Enregistrements Controle d'entree".
# Sert les fichiers de l'application sur http://localhost (votre PC uniquement, rien ne sort sur internet),
# pour que le navigateur autorise l'enregistrement direct dans le classeur Excel.
# Lance par "Lancer l'application.bat". Fermer la fenetre arrete le serveur.

$ErrorActionPreference = 'Stop'
$racine = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $racine.EndsWith('\')) { $racine = $racine + '\' }

$types = @{
  '.html' = 'text/html; charset=utf-8'
  '.js'   = 'text/javascript; charset=utf-8'
  '.css'  = 'text/css; charset=utf-8'
  '.json' = 'application/json; charset=utf-8'
  '.svg'  = 'image/svg+xml'
  '.png'  = 'image/png'
  '.ico'  = 'image/x-icon'
  '.md'   = 'text/plain; charset=utf-8'
}

# premier port libre entre 8765 et 8775 (adresse locale uniquement : pas besoin de droits administrateur)
$serveur = $null
$port = 0
foreach ($p in 8765..8775) {
  try {
    $l = New-Object System.Net.HttpListener
    $l.Prefixes.Add("http://localhost:$p/")
    $l.Start()
    $serveur = $l
    $port = $p
    break
  } catch { }
}
if ($null -eq $serveur) {
  Write-Host ''
  Write-Host 'Impossible de demarrer le serveur local (ports 8765 a 8775 occupes ou bloques).'
  Write-Host 'Utilisez la version en ligne : https://leomazet0809-a11y.github.io/Enreg-Ctrl/'
  exit 1
}

$adresse = "http://localhost:$port/index.html"
Write-Host ''
Write-Host '  Enregistrements Controle d''entree'
Write-Host "  Application ouverte a l'adresse : $adresse"
Write-Host ''
Write-Host '  Laissez cette fenetre ouverte (vous pouvez la reduire) pendant l''utilisation.'
Write-Host '  Fermez-la pour arreter l''application.'
Write-Host ''
Start-Process $adresse

while ($serveur.IsListening) {
  try {
    $ctx = $serveur.GetContext()
    $rep = $ctx.Response
    try {
      $chemin = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath.TrimStart('/'))
      if ([string]::IsNullOrEmpty($chemin)) { $chemin = 'index.html' }
      $fichier = [IO.Path]::GetFullPath((Join-Path $racine $chemin))
      $dansDossier = $fichier.StartsWith($racine, [StringComparison]::OrdinalIgnoreCase)
      if ($dansDossier -and (Test-Path -LiteralPath $fichier -PathType Leaf)) {
        $ext = [IO.Path]::GetExtension($fichier).ToLowerInvariant()
        if ($types.ContainsKey($ext)) { $rep.ContentType = $types[$ext] } else { $rep.ContentType = 'application/octet-stream' }
        $rep.Headers.Add('Cache-Control', 'no-cache')
        $octets = [IO.File]::ReadAllBytes($fichier)
        $rep.ContentLength64 = $octets.Length
        $rep.OutputStream.Write($octets, 0, $octets.Length)
      } else {
        $rep.StatusCode = 404
      }
    } finally {
      $rep.OutputStream.Close()
    }
  } catch { }
}
