# Unauthenticated localhost control service

A writeup of a flaw in the first version of this project's bridge server, found
before release. The vulnerable file is kept in the repo as
`rehearsal-bridge.v1.ps1` so the two versions can be read side by side.

**Severity:** low impact, but a common and frequently underestimated class.
**Status:** fixed before first release.

---

## Summary

`rehearsal-bridge.v1.ps1` ran an HTTP server on `127.0.0.1:8765` that drove a live
PowerPoint slide show. It had no authentication, answered `GET` requests with side
effects, and returned `Access-Control-Allow-Origin: *`.

While it was running, **any web page open in any tab of the same browser could
control the slide show.**

---

## The vulnerable design

Three decisions combined to create the problem:

```powershell
# rehearsal-bridge.v1.ps1

$path = ($req -split ' ')[1]              # 1. no credential of any kind

if ($path -like "/goto*") {               # 2. GET performs the action
    $n = 1
    if ($path -match 'n=(\d+)') { $n = [int]$Matches[1] }
    $ok = Go-Slide $n
}

$head = "HTTP/1.1 200 OK`r`n" +
        "Access-Control-Allow-Origin: *`r`n"   # 3. every origin trusted
```

Individually each looks harmless on a loopback service. Together they are a
remotely reachable control channel.

## Proof of concept

Any page on any domain, with the bridge running:

```html
<!-- Full control. No fetch, no JavaScript required for the side effect. -->
<img src="http://127.0.0.1:8765/goto?n=1" hidden>

<script>
  // Enumerate whether the bridge is running, and in which mode.
  fetch("http://127.0.0.1:8765/health")
    .then(r => r.json())
    .then(j => console.log("bridge is up:", j.mode));
</script>
```

The `<img>` tag is the part worth sitting with. It is not blocked by the same-origin
policy, needs no scripting, and works from any site.

## Why the CORS header was not the real problem

The instinct is that `Access-Control-Allow-Origin: *` was the bug. It made things
worse, but removing it would not have fixed anything.

**CORS governs whether a page may read a response. It does not govern whether the
request is sent.** For a "simple" request — a `GET`, or a `POST` with an ordinary
content type — the browser delivers it to the server and the server acts on it.
The response is then withheld from the calling page. The side effect has already
happened.

So with `ACAO: *` removed, an attacker still advances your slides. They just cannot
read the confirmation.

What `ACAO: *` *did* add was reconnaissance: any site could read `/health` and learn
that the bridge was running and which mode it was in.

## Why localhost is not a boundary

Two assumptions were wrong.

**"Only software on this machine can reach it."** The browser is software on this
machine, and it executes code from every site you visit. Binding to loopback keeps
the port off the network; it does nothing about the browser.

**"An attacker would have to know it exists."** Scanning a handful of common ports
from JavaScript takes milliseconds. Localhost services on predictable ports have
produced a long series of real CVEs in developer tools, media servers and desktop
agents.

There is also **DNS rebinding**, which turns "local only" into "any website".
An attacker serves a page from `evil.example` with a very short DNS TTL, then
re-answers with `127.0.0.1`. The browser now believes `evil.example:8765` and the
bridge are the same origin, and hands over full same-origin access — reads included.

## Impact

Realistic worst case for the intended setup: someone advances or blanks your slides
during a talk. Embarrassing, not dangerous.

Two things raise it above trivial:

- **Keystroke mode is worse than COM mode.** With `-Target`, the bridge activates a
  window by title and sends keys. If the title does not match, `AppActivate` fails
  and the keystroke lands in whatever window currently has focus. An attacker
  triggering `/goto` repeatedly is injecting keystrokes into an arbitrary
  application.
- **The pattern generalises badly.** The same three decisions in a service that
  touches files, credentials or a shell is a serious vulnerability. The mistake is
  the design, not the blast radius.

---

## The fix

Four controls, in `rehearsal-bridge.ps1`.

### 1. Per-session bearer token

192 bits from a CSPRNG, minted at startup, required on every command in the
`X-Rehearsal-Token` header. Compared in constant time.

```powershell
$bytes = New-Object byte[] 24
(New-Object System.Security.Cryptography.RNGCryptoServiceProvider).GetBytes($bytes)
$TOKEN = -join ($bytes | ForEach-Object { $_.ToString("x2") })
```

The token is delivered by having the bridge launch the page itself with
`#t=<token>` appended. **URL fragments are never sent to a server**, so the
credential does not travel over the wire, and the page clears it from the address
bar with `history.replaceState` once read.

### 2. A custom header, which forces a preflight

This is the control that actually closes the hole, and it works because of a
property of CORS rather than in spite of one.

`X-Rehearsal-Token` is not a CORS-safelisted request header, so a browser will not
send it on a cross-origin request without first making an `OPTIONS` preflight and
receiving explicit permission. The bridge only grants that permission to allowed
origins.

A hostile page therefore cannot attach the header at all — so it can no longer
reach the endpoint even though the endpoint is on the same machine, and even
though it could still fire a bare `GET`. A bare `GET` is now rejected as `405`.

### 3. Host header validation

```powershell
function Test-HostHeader([string]$h) {
  $name = ($h -split ':')[0]
  return ($name -eq "127.0.0.1" -or $name -eq "localhost")
}
```

This is what stops DNS rebinding. The browser sends `Host: evil.example` because
that is what the user typed, regardless of what it resolved to. The bridge only
answers to its own names.

### 4. Loopback binding, asserted

Bound to `IPAddress.Loopback`, and the remote address is checked on every
connection rather than assumed from the bind.

### Also

- **POST only.** Side effects do not belong on a method an `<img>` tag can trigger.
- **Bounds checking** on the slide number before it reaches COM.
- **Idle shutdown** after 30 minutes, so a forgotten console window is not a
  standing service.
- **Rejections are logged** to the console with the offending origin or host.

---

## What each control stops

| Attack | Stopped by |
|---|---|
| `<img src>` from any site | POST-only |
| `fetch` from any site | Custom header forces preflight; token |
| Reading `/health` to fingerprint | Origin allowlist; token |
| DNS rebinding | Host header validation |
| Reaching it from the network | Loopback bind + remote address assertion |
| Guessing the token | 192 bits of CSPRNG output |
| Token recovered from server logs or history | Fragment delivery; fragment stripped after read |

## What this does not protect against

Worth stating plainly, because a writeup that claims total coverage is not credible.

- **Any process already running as your user.** It can read the console, the
  browser's memory, or simply drive PowerPoint through COM itself. If an attacker
  is executing code as you, this service is not the weak point.
- **The `null` origin is weak.** A page loaded from `file://` sends `Origin: null`,
  and a sandboxed iframe can produce `null` too, so the allowlist cannot fully
  distinguish them. This is precisely why the token — not the origin check — is the
  primary control. Serving the page over `http://localhost` instead of opening it
  from disk would produce a real origin and tighten this.
- **The token appears in the console window** and briefly in the browser's address
  bar. Anyone who can see your screen can copy it. Acceptable here; not acceptable
  for anything that matters.
- **No rate limiting.** Irrelevant against a 192-bit token, but it would matter if
  the credential were ever weakened.
- **No TLS.** Deliberate: loopback traffic does not leave the machine, and a
  self-signed certificate on `127.0.0.1` would train the user to click through
  certificate warnings, which is a worse outcome.

## Lessons

1. **A local service is a service.** The threat model that matters is not "who is
   on my network" but "what is my browser willing to send on behalf of a stranger".
2. **CORS is not access control.** It decides who may read a response, not who may
   cause an effect. Anything that changes state needs its own credential.
3. **Requiring a custom header is a real control,** because it forces a preflight the
   server can refuse. It is a small change that removes an entire attack path.
4. **Validate `Host`,** or loopback binding is a weaker guarantee than it looks.
5. The bug came from thinking of `127.0.0.1` as a trust boundary. It is an address.

## Timeline

| | |
|---|---|
| v1 written | Working bridge, no authentication |
| Review | Flaw identified during a security pass over my own code, before release |
| v2 | Four controls added; v1 retained in-repo for this writeup |
| Published | Both versions, first public release |
