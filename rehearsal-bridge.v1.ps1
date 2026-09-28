<#
  rehearsal-bridge.ps1
  Lets the rehearsal page drive a real PowerPoint slide show, animations and all.

  It opens the deck, starts the show, then listens on 127.0.0.1:8765 for commands
  from the browser. PowerPoint is driven through its automation interface rather
  than simulated keystrokes, so keyboard focus never leaves your rehearsal window
  and the clicker keeps working.

  Run it with Start-Rehearsal-Bridge.cmd, or:
     powershell -ExecutionPolicy Bypass -File rehearsal-bridge.ps1

  Ctrl+C to stop.
#>

param(
  [string]$Deck = "",
  [int]$Port = 8765,
  [string]$Target = ""   # optional: window title to SendKeys instead (LibreOffice Impress etc.)
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

function Say($msg, $color = "Gray") { Write-Host $msg -ForegroundColor $color }

Say ""
Say "  Hashing rehearsal bridge" "Cyan"
Say "  ------------------------" "DarkGray"

# ---------------------------------------------------------------- find the deck
if (-not $Target) {
  if (-not $Deck) {
    $found = Get-ChildItem -Path $here -Filter *.pptx -File -ErrorAction SilentlyContinue |
             Sort-Object Name | Select-Object -First 1
    if ($found) { $Deck = $found.FullName }
  }
  if (-not $Deck -or -not (Test-Path $Deck)) {
    Say "  No .pptx found next to this script." "Yellow"
    Say "  Put hashing-iot.pptx in the same folder, or run:" "DarkGray"
    Say "     powershell -ExecutionPolicy Bypass -File rehearsal-bridge.ps1 -Deck ""C:\path\to\deck.pptx""" "DarkGray"
    Read-Host "`n  Press Enter to close"
    exit 1
  }
}

# ------------------------------------------------------------ connect to viewer
$ppt = $null
$mode = "sendkeys"

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
    Say "  Falling back to keystrokes. Pass -Target to name the window, e.g. -Target Impress" "DarkGray"
    $ppt = $null
    $mode = "sendkeys"
  }
}

$wsh = New-Object -ComObject WScript.Shell

# Win32 focus save/restore - AppActivate has to guess at window titles, this does not.
if (-not ("Win32Focus" -as [type])) {
  Add-Type -Namespace Native -Name Win32Focus -MemberDefinition @"
    [DllImport("user32.dll")] public static extern System.IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr hWnd);
"@ -ErrorAction SilentlyContinue
}
function Invoke-Key([string]$keys) {
  $prev = [System.IntPtr]::Zero
  try { $prev = [Native.Win32Focus]::GetForegroundWindow() } catch { }
  $wsh.AppActivate($Target) | Out-Null
  Start-Sleep -Milliseconds 70
  $wsh.SendKeys($keys)
  Start-Sleep -Milliseconds 45
  try { if ($prev -ne [System.IntPtr]::Zero) { [Native.Win32Focus]::SetForegroundWindow($prev) | Out-Null } }
  catch { $wsh.AppActivate("Hashing") | Out-Null }
}
if ($mode -eq "sendkeys") {
  if (-not $Target) { $Target = "Impress" }
  Say "  Keystroke mode - will send keys to a window matching '$Target'." "Yellow"
  Say "  Focus will flicker on each click. COM mode is smoother if you have PowerPoint." "DarkGray"
}

function Get-View {
  try { if ($ppt.SlideShowWindows.Count -ge 1) { return $ppt.SlideShowWindows.Item(1).View } } catch { }
  return $null
}

function Go-Slide([int]$n) {
  if ($mode -eq "com") {
    $v = Get-View
    if ($null -eq $v) { $pres.SlideShowSettings.Run() | Out-Null; Start-Sleep -Milliseconds 600; $v = Get-View }
    if ($null -ne $v) { $v.GotoSlide($n) | Out-Null; return $true }
    return $false
  }
  # keystroke fallback: focus the viewer, send a key, hand focus straight back
  try { Invoke-Key "{PGDN}"; return $true } catch { return $false }
}

function Set-Blank([bool]$on) {
  if ($mode -eq "com") {
    $v = Get-View
    if ($null -ne $v) { $v.State = $(if ($on) { 3 } else { 1 }); return $true }   # 3 = black, 1 = running
    return $false
  }
  try { Invoke-Key "b"; return $true } catch { return $false }
}

# --------------------------------------------------------------- tiny HTTP server
# TcpListener rather than HttpListener: no admin rights and no URL reservation needed.
$listener = $null
foreach ($p in @($Port, $Port + 1, $Port + 2)) {
  try {
    $try = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $p)
    $try.Start(); $listener = $try; $Port = $p; break
  } catch { }
}
if (-not $listener) {
  Say "  Ports $Port-$($Port+2) are all busy. Close any other bridge window and retry." "Red"
  Read-Host "`n  Press Enter to close"; exit 1
}

Say ""
Say "  Listening on http://127.0.0.1:$Port" "Cyan"
Say "  Now open hashing-rehearsal.html - it connects on its own." "White"
Say "  Leave this window open. Ctrl+C to stop." "DarkGray"
Say ""

function Send-Reply($client, $bodyText) {
  $bytes = [Text.Encoding]::UTF8.GetBytes($bodyText)
  $head = "HTTP/1.1 200 OK`r`n" +
          "Access-Control-Allow-Origin: *`r`n" +
          "Content-Type: application/json`r`n" +
          "Cache-Control: no-store`r`n" +
          "Content-Length: $($bytes.Length)`r`n" +
          "Connection: close`r`n`r`n"
  $out = $client.GetStream()
  $hb = [Text.Encoding]::ASCII.GetBytes($head)
  $out.Write($hb, 0, $hb.Length)
  $out.Write($bytes, 0, $bytes.Length)
  $out.Flush()
}

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $reader = New-Object System.IO.StreamReader($client.GetStream())
      $req = $reader.ReadLine()
      if (-not $req) { $client.Close(); continue }

      $path = ($req -split ' ')[1]
      $ok = $true
      $note = ""

      if ($path -like "/goto*") {
        $n = 1
        if ($path -match 'n=(\d+)') { $n = [int]$Matches[1] }
        $ok = Go-Slide $n
        $note = "slide $n"
        Write-Host ("  -> slide {0}" -f $n) -ForegroundColor DarkGray
      }
      elseif ($path -like "/blank*")   { $ok = Set-Blank $true;  $note = "black" }
      elseif ($path -like "/unblank*") { $ok = Set-Blank $false; $note = "running" }
      elseif ($path -like "/health*")  { $note = $mode }
      else { $note = "ignored" }

      Send-Reply $client ('{"ok":' + $ok.ToString().ToLower() + ',"mode":"' + $mode + '","note":"' + $note + '"}')
    } catch {
      Write-Host "  ! $($_.Exception.Message)" -ForegroundColor DarkYellow
    } finally {
      $client.Close()
    }
  }
} finally {
  $listener.Stop()
  Say "`n  Bridge stopped." "DarkGray"
}
