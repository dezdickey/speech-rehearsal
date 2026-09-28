<#
  rehearsal-bridge.ps1
  Drives a real PowerPoint slide show from the rehearsal page, so animations
  play and keyboard focus never leaves the browser.

  SECURITY MODEL
  --------------
  This is an HTTP service on your own machine, and a browser will happily let
  any page you visit talk to 127.0.0.1. The first version of this script trusted
  every origin and required no credential, which meant any site open in another
  tab could drive the slide show. Four controls close that:

    1. Bearer token   192 random bits, minted per run, required in the
                      X-Rehearsal-Token header on every command.
    2. Custom header  Requiring a non-standard header forces the browser to send
                      a CORS preflight, which we answer only for known origins.
                      A hostile page cannot attach that header at all.
    3. Host check     Requests must be addressed to 127.0.0.1 or localhost. This
                      is what defeats DNS rebinding, where an attacker's domain
                      resolves to loopback to inherit its privileges.
    4. Loopback bind  The socket binds to 127.0.0.1 explicitly and the remote
                      address is asserted, so it is never reachable off-machine.

  Commands are POST, not GET: side effects do not belong on a method that an
  <img> tag can trigger.

  Run with Start-Rehearsal-Bridge.cmd, or:
     powershell -ExecutionPolicy Bypass -File rehearsal-bridge.ps1

  Ctrl+C to stop. It also stops itself after 30 idle minutes.
#>

param(
  [string]$Deck = "",
  [int]$Port = 8765,
  [string]$Target = "",           # window title for keystroke mode (LibreOffice Impress)
  [int]$IdleMinutes = 30,
  [switch]$NoBrowser
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
function Say($m, $c = "Gray") { Write-Host $m -ForegroundColor $c }

Say ""
Say "  Speech rehearsal bridge" "Cyan"
Say "  -----------------------" "DarkGray"

# ------------------------------------------------------------------ credential
$bytes = New-Object byte[] 24
(New-Object System.Security.Cryptography.RNGCryptoServiceProvider).GetBytes($bytes)
$TOKEN = -join ($bytes | ForEach-Object { $_.ToString("x2") })

# Constant-time compare. Overkill at this scale, but ordinary string equality on
# a secret leaks length and prefix through timing, and the habit is worth keeping.
function Test-Token([string]$given) {
  if ([string]::IsNullOrEmpty($given)) { return $false }
  if ($given.Length -ne $TOKEN.Length) { return $false }
  $diff = 0
  for ($i = 0; $i -lt $TOKEN.Length; $i++) {
    $diff = $diff -bor ([int][char]$TOKEN[$i] -bxor [int][char]$given[$i])
  }
  return ($diff -eq 0)
}

# Origins we will answer a preflight for. A page loaded from disk sends "null" -
# that is Chrome's origin for file://, and it is genuinely weak, because a
# sandboxed iframe can produce it too. The token is what actually guards the
# endpoint; this list only limits who may ask.
$ALLOWED_ORIGINS = @("null", "http://localhost", "http://127.0.0.1")
function Test-Origin([string]$o) {
  if ([string]::IsNullOrEmpty($o)) { return $true }          # non-browser client
  foreach ($a in $ALLOWED_ORIGINS) { if ($o -eq $a -or $o.StartsWith($a + ":")) { return $true } }
  return $false
}
function Test-HostHeader([string]$h) {
  if ([string]::IsNullOrEmpty($h)) { return $false }
  $name = ($h -split ':')[0]
  return ($name -eq "127.0.0.1" -or $name -eq "localhost")
}

# ---------------------------------------------------------------- find the deck
if (-not $Target) {
  if (-not $Deck) {
    $f = Get-ChildItem -Path $here -Filter *.pptx -File -ErrorAction SilentlyContinue |
         Sort-Object Name | Select-Object -First 1
    if ($f) { $Deck = $f.FullName }
  }
  if (-not $Deck -or -not (Test-Path $Deck)) {
    Say "  No .pptx found next to this script." "Yellow"
    Say "  Put a .pptx in this folder (demo-deck.pptx ships with it), or pass -Deck ""C:\path\deck.pptx""" "DarkGray"
    Read-Host "`n  Press Enter to close"; exit 1
  }
}

# ------------------------------------------------------------ connect to viewer
$ppt = $null; $pres = $null; $mode = "sendkeys"
if (-not $Target) {
  try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $true
    Say "  Opening $(Split-Path -Leaf $Deck) ..." "DarkGray"
    $pres = $ppt.Presentations.Open($Deck, $false, $false, $true)
    $pres.SlideShowSettings.Run() | Out-Null
    Start-Sleep -Milliseconds 900
    $mode = "com"
    Say "  PowerPoint slide show running - animations live." "Green"
  } catch {
    Say "  Could not drive PowerPoint ($($_.Exception.Message))." "Yellow"
    $ppt = $null; $mode = "sendkeys"
  }
}

# ------------------------------------------------ or LibreOffice Impress, directly
# LibreOffice registers an automation (COM) bridge on Windows. It drives the show the
# way COM drives PowerPoint: no keystrokes, no focus changes. Keystrokes can't work
# reliably with Impress anyway: its show window can't take keyboard focus.
# PowerShell's usual COM binding needs type information this bridge doesn't provide,
# so every call goes through InvokeMember. Only tried when LibreOffice is already
# running, so the bridge never starts it in the background.
$uno = $null; $unoCtl = $null
function Uno($o, [string]$m, [object[]]$a = @()) {
  [System.__ComObject].InvokeMember($m, [Reflection.BindingFlags]::InvokeMethod, $null, $o, $a)
}
function Get-ImpressShow {
  # The running show's controller. Starts the show if a presentation is open but not
  # showing. Looked up again whenever the cached one has stopped (show ended, restarted).
  try { if ($script:unoCtl -and (Uno $script:unoCtl "isRunning")) { return $script:unoCtl } } catch { }
  $script:unoCtl = $null
  try {
    $desk = Uno $uno "createInstance" @("com.sun.star.frame.Desktop")
    $en = Uno (Uno $desk "getComponents") "createEnumeration"; $idle = $null
    while (Uno $en "hasMoreElements") {
      $c = Uno $en "nextElement"
      if (-not (Uno $c "supportsService" @("com.sun.star.presentation.PresentationDocument"))) { continue }
      $pr = Uno $c "getPresentation"
      if (Uno $pr "isRunning") { $script:unoCtl = Uno $pr "getController"; return $script:unoCtl }
      if (-not $idle) { $idle = $pr }
    }
    if ($idle) {
      Uno $idle "start" | Out-Null
      for ($k = 0; $k -lt 20 -and -not $script:unoCtl; $k++) { Start-Sleep -Milliseconds 150; $script:unoCtl = Uno $idle "getController" }
    }
  } catch { $script:unoCtl = $null }
  return $script:unoCtl
}
if ($mode -eq "sendkeys" -and -not $Target -and (Get-Process -Name soffice.bin -ErrorAction SilentlyContinue)) {
  try {
    $uno = New-Object -ComObject com.sun.star.ServiceManager
    if (Get-ImpressShow) { $mode = "uno"; Say "  LibreOffice Impress slide show found - driving it directly, no keystrokes." "Green" }
    else { Say "  LibreOffice is running, but no presentation is open in it." "Yellow" }
  } catch { Say "  Could not drive LibreOffice Impress ($($_.Exception.Message))." "Yellow" }
}

if ($mode -eq "sendkeys") {
  # Last resort: keystrokes, which need window switching. Only set up when used.
  $wsh = New-Object -ComObject WScript.Shell
  if (-not ("Native.Win32Focus" -as [type])) {
    Add-Type -Namespace Native -Name Win32Focus -MemberDefinition @"
    [DllImport("user32.dll")] public static extern System.IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr hWnd);
"@ -ErrorAction SilentlyContinue
  }
  if (-not $Target) { $Target = "Impress" }
  Say "  Keystroke mode - sending keys to a window matching '$Target'." "Yellow"
}

function Invoke-Key([string]$keys) {
  $prev = [System.IntPtr]::Zero
  try { $prev = [Native.Win32Focus]::GetForegroundWindow() } catch { }
  $wsh.AppActivate($Target) | Out-Null
  Start-Sleep -Milliseconds 70
  $wsh.SendKeys($keys)
  Start-Sleep -Milliseconds 45
  try { if ($prev -ne [System.IntPtr]::Zero) { [Native.Win32Focus]::SetForegroundWindow($prev) | Out-Null } } catch { }
}
function Get-View {
  try { if ($ppt.SlideShowWindows.Count -ge 1) { return $ppt.SlideShowWindows.Item(1).View } } catch { }
  return $null
}
function Go-Slide([int]$n) {
  if ($n -lt 1 -or $n -gt 999) { return $false }             # never trust the wire
  if ($mode -eq "com") {
    $v = Get-View
    if ($null -eq $v) { $pres.SlideShowSettings.Run() | Out-Null; Start-Sleep -Milliseconds 600; $v = Get-View }
    if ($null -ne $v) { $v.GotoSlide($n) | Out-Null; return $true }
    return $false
  }
  if ($mode -eq "uno") {
    try {
      $c = Get-ImpressShow; if ($null -eq $c) { return $false }
      $last = [int](Uno $c "getSlideCount")
      Uno $c "gotoSlideIndex" @([Math]::Min($n, $last) - 1) | Out-Null   # a jump, so Back works too
      return $true
    } catch { return $false }
  }
  try { Invoke-Key "{PGDN}"; return $true } catch { return $false }
}
function Set-Blank([bool]$on) {
  if ($mode -eq "com") {
    $v = Get-View
    if ($null -ne $v) { $v.State = $(if ($on) { 3 } else { 1 }); return $true }
    return $false
  }
  if ($mode -eq "uno") {
    try {
      $c = Get-ImpressShow; if ($null -eq $c) { return $false }
      if ($on) { Uno $c "blankScreen" @(0) | Out-Null } else { Uno $c "resume" | Out-Null }
      return $true
    } catch { return $false }
  }
  try { Invoke-Key "b"; return $true } catch { return $false }
}

# ------------------------------------------------------------- socket, loopback
$listener = $null
foreach ($p in @($Port, ($Port + 1), ($Port + 2))) {     # parentheses: "," binds tighter than "+"
  try {
    $t = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $p)
    $t.Start(); $listener = $t; $Port = $p; break
  } catch { }
}
if (-not $listener) {
  Say "  Ports $Port-$($Port+2) are busy. Close any other bridge and retry." "Red"
  Read-Host "`n  Press Enter to close"; exit 1
}

# The token rides in the URL fragment, which browsers never put on the wire.
# Start-Process on a file:// URL goes through the file association, and Windows drops
# the #fragment on the way, so the page would open without its token. Hand the URL to
# the default browser's own program instead, as an argument. If that can't be found,
# fall back to the association; the page then asks for the token.
function Open-Page([string]$url) {
  $exe = $null
  try {
    $prog = (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice" -ErrorAction Stop).ProgId
    $cmd = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$prog\shell\open\command" -ErrorAction Stop).'(default)'
    if ($cmd -match '^\s*"([^"]+\.exe)"' -or $cmd -match '^\s*(\S+\.exe)') { $exe = $Matches[1] }
  } catch { }
  if ($exe -and (Test-Path -LiteralPath $exe)) { Start-Process -FilePath $exe -ArgumentList ('"' + $url + '"') | Out-Null }
  else { Start-Process $url | Out-Null }
}
$pageFile = Join-Path $here "speech-rehearsal.html"
$pageUrl  = "file:///" + ($pageFile -replace '\\', '/') + "#t=$TOKEN"

Say ""
Say "  Listening on http://127.0.0.1:$Port  (loopback only)" "Cyan"
Say "  Session token: $TOKEN" "DarkGray"
Say ""
if ((Test-Path $pageFile) -and -not $NoBrowser) {
  Say "  Opening the rehearsal page with this session's token..." "White"
  try { Open-Page $pageUrl }
  catch { Say "  Could not launch a browser. Open this URL yourself:" "Yellow"; Say "  $pageUrl" "DarkGray" }
} else {
  Say "  Open speech-rehearsal.html and paste the token above when asked." "White"
}
Say ""
Say "  Leave this window open. Ctrl+C to stop; it also stops after $IdleMinutes idle minutes." "DarkGray"
Say ""

function Send-Reply($client, $status, $bodyText, $extraHeaders) {
  $b = [Text.Encoding]::UTF8.GetBytes($bodyText)
  $h = "HTTP/1.1 $status`r`n" +
       "Content-Type: application/json`r`n" +
       "Cache-Control: no-store`r`n" +
       "X-Content-Type-Options: nosniff`r`n" +
       $extraHeaders +
       "Content-Length: $($b.Length)`r`n" +
       "Connection: close`r`n`r`n"
  $s = $client.GetStream()
  $hb = [Text.Encoding]::ASCII.GetBytes($h)
  $s.Write($hb, 0, $hb.Length); $s.Write($b, 0, $b.Length); $s.Flush()
}

$lastSeen = Get-Date
try {
  while ($true) {
    if (-not $listener.Pending()) {
      Start-Sleep -Milliseconds 120
      if (((Get-Date) - $lastSeen).TotalMinutes -ge $IdleMinutes) {
        Say "  Idle for $IdleMinutes minutes - shutting down." "DarkGray"; break
      }
      continue
    }
    $client = $listener.AcceptTcpClient()
    try {
      # Bound to loopback already, but assert it rather than assume it.
      $remote = $client.Client.RemoteEndPoint.Address.ToString()
      if ($remote -ne "127.0.0.1" -and $remote -ne "::1") { $client.Close(); continue }

      $reader  = New-Object System.IO.StreamReader($client.GetStream())
      $reqLine = $reader.ReadLine()
      if (-not $reqLine) { $client.Close(); continue }

      $hdr = @{}
      while ($true) {
        $line = $reader.ReadLine()
        if ($null -eq $line -or $line -eq "") { break }
        $ix = $line.IndexOf(':')
        if ($ix -gt 0) { $hdr[$line.Substring(0, $ix).Trim().ToLower()] = $line.Substring($ix + 1).Trim() }
      }

      $parts  = $reqLine -split ' '
      $method = $parts[0]
      $path   = $parts[1]
      $origin = $hdr["origin"]
      $acao   = if ($origin) { "Access-Control-Allow-Origin: $origin`r`nVary: Origin`r`n" } else { "" }

      if (-not (Test-HostHeader $hdr["host"])) {
        # DNS rebinding: an attacker domain pointed at 127.0.0.1 dies here.
        Send-Reply $client "403 Forbidden" '{"ok":false,"error":"bad host"}' ""
        Write-Host "  ! rejected host: $($hdr["host"])" -ForegroundColor DarkYellow
        $client.Close(); continue
      }
      if (-not (Test-Origin $origin)) {
        Send-Reply $client "403 Forbidden" '{"ok":false,"error":"origin not allowed"}' ""
        Write-Host "  ! rejected origin: $origin" -ForegroundColor DarkYellow
        $client.Close(); continue
      }

      if ($method -eq "OPTIONS") {
        $pre = $acao +
               "Access-Control-Allow-Methods: POST, OPTIONS`r`n" +
               "Access-Control-Allow-Headers: X-Rehearsal-Token`r`n" +
               "Access-Control-Max-Age: 600`r`n"
        Send-Reply $client "204 No Content" "" $pre
        $client.Close(); continue
      }
      if ($method -ne "POST") {
        Send-Reply $client "405 Method Not Allowed" '{"ok":false,"error":"use POST"}' $acao
        $client.Close(); continue
      }
      if (-not (Test-Token $hdr["x-rehearsal-token"])) {
        Send-Reply $client "401 Unauthorized" '{"ok":false,"error":"bad or missing token"}' $acao
        Write-Host "  ! rejected: bad token from origin '$origin'" -ForegroundColor DarkYellow
        $client.Close(); continue
      }

      $lastSeen = Get-Date
      $ok = $true; $note = ""
      if ($path -like "/goto*") {
        $n = 1; if ($path -match 'n=(\d+)') { $n = [int]$Matches[1] }
        $ok = Go-Slide $n; $note = "slide $n"
        Write-Host ("  -> slide {0}" -f $n) -ForegroundColor DarkGray
      }
      elseif ($path -like "/blank*")   { $ok = Set-Blank $true;  $note = "black" }
      elseif ($path -like "/unblank*") { $ok = Set-Blank $false; $note = "running" }
      elseif ($path -like "/health*")  { $note = $mode }
      else { $ok = $false; $note = "unknown command" }

      Send-Reply $client "200 OK" ('{"ok":' + $ok.ToString().ToLower() + ',"mode":"' + $mode + '","note":"' + $note + '"}') $acao
    } catch {
      Write-Host "  ! $($_.Exception.Message)" -ForegroundColor DarkYellow
    } finally {
      try { $client.Close() } catch { }
    }
  }
} finally {
  $listener.Stop()
  Say "`n  Bridge stopped. This session's token is now dead." "DarkGray"
}
