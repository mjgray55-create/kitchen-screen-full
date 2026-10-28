# Builds test/test.html: the real Code.gs running on a pretend Google Sheet, plus the real Index.html.
# Run:  powershell -ExecutionPolicy Bypass -File test\build.ps1     then serve the test folder (serve.ps1).
$root = Split-Path $PSScriptRoot -Parent
$code = [IO.File]::ReadAllText("$root\Code.gs")
$fake = [IO.File]::ReadAllText("$PSScriptRoot\fake-sheets.js")
$page = [IO.File]::ReadAllText("$root\Index.html")
$server = @"
<script>
window.__ERRORS__ = [];
window.addEventListener('error', function (e) { window.__ERRORS__.push(String(e.message)); });
$fake
</script>
<script>
// Code.gs inside its own scope, so its names can't clash with the screen's
var SERVER = (function () {
$code
  return function (fn, args) {
    if (/_$/.test(fn) || typeof eval(fn) !== 'function') throw new Error('No server function ' + fn);
    return eval(fn).apply(null, args);
  };
})();
// like google.script.run: arguments and results go through JSON
window.__MOCK__ = new Proxy({}, { get: function (t, fn) { return function () {
  var args = JSON.parse(JSON.stringify(Array.prototype.slice.call(arguments)));
  var out = SERVER(fn, args);
  return out === undefined ? undefined : JSON.parse(JSON.stringify(out));
}; } });
</script>
<meta name="viewport" content="width=device-width, initial-scale=1">
"@
$out = $page.Replace('<base target="_top">', $server)
[IO.File]::WriteAllText("$PSScriptRoot\test.html", $out, (New-Object Text.UTF8Encoding $false))
"built test.html"
