# Serves the test folder at http://localhost:8765/test.html (Ctrl+C to stop).
$root = $PSScriptRoot
$l = New-Object Net.HttpListener
$l.Prefixes.Add('http://localhost:8765/')
$l.Start()
Write-Host 'serving on http://localhost:8765/test.html'
while ($l.IsListening) {
  $c = $l.GetContext()
  $p = Join-Path $root ($c.Request.Url.LocalPath.TrimStart('/'))
  if (Test-Path $p -PathType Leaf) {
    $b = [IO.File]::ReadAllBytes($p)
    $c.Response.ContentType = 'text/html; charset=utf-8'
    $c.Response.OutputStream.Write($b, 0, $b.Length)
  } else { $c.Response.StatusCode = 404 }
  $c.Response.Close()
}
