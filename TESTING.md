# Testing

Roughly ten minutes. Do the phases in order — Phase A decides whether the rest is
worth running at all.

Everything in one folder: `rehearsal-bridge.ps1`, `rehearsal-bridge.v1.ps1`,
`Start Speech Rehearsal.cmd`, `speech-rehearsal.html`, `poc.html`, `demo-deck.pptx`.

## Before you start

**Unblock the downloaded scripts.** Windows tags files from the internet and
PowerShell will refuse to run them. In PowerShell, in that folder:

```powershell
Get-ChildItem *.ps1 | Unblock-File
```

**Your antivirus may block the bridge.** On the author's PC, Bitdefender stopped the
bridge the first time it was run, and later quarantined `rehearsal-bridge.ps1`, with a
heuristic detection: `CMD:Heur.BZC.PZQ.Boxter.1196.BA2C9B7D`. A local web server,
Office automation and a keystroke fallback all resemble patterns malware uses. Expect
the same for `rehearsal-bridge.v1.ps1`, the deliberately vulnerable version, which is
only for Phase B.
- If it happens, write down the product, the detection name and the time. That is a
  result worth recording, not a reason to abandon the test.
- Restoring the file and adding an exception is your decision. An exception usually
  covers one exact path, so don't move or rename the file afterwards, and run the
  tests from that folder.
- Don't upload the files to VirusTotal: it's owned by Google.

**The app starts with no speech.** When it opens, click **Try the sample**: its 4
slides match `demo-deck.pptx`. Or paste or open your own speech, marked up for your own
deck. The security phases use `poc.html` and `curl.exe`, not the app, so they don't
need this.

**If a firewall prompt appears, deny it.** The bridge binds to loopback only and
does not need any firewall exception. If denying it breaks the bridge, that is
itself a finding worth writing down.

---


## Two sessions

The security controls do not depend on which presentation app is running. Reading
the bridge source: the Host check, Origin check, method check and token check all
return **before** `Go-Slide` is ever called. So the whole test suite works with
LibreOffice today; only the COM automation test needs PowerPoint.

| Session | Needs | Covers |
|---|---|---|
| **1 — today** | LibreOffice | Phase L, then B, C, D, E — every access control |
| **2 — tomorrow** | PowerPoint | Phase A — COM automation, the architecture check |

---

## Phase L — LibreOffice (today, 5 min)

Install LibreOffice from the official site. Either branch works; "Still" is the
more conservative one.

> https://www.libreoffice.org/download/download-libreoffice/

### L1 — does the deck play in Impress?

1. Open `demo-deck.pptx` in Impress and start the show: **Slide Show → Start from
   First Slide**. (F5 works too, unless your keyboard's function-key layer takes it;
   then use Fn+F5.)
2. Click through all four slides, one click each: "Street trees", "The problem",
   "What trees do", "The ask".

The demo deck has no animations. If you test with your own animated deck instead,
check that each slide's animations play on their own and nothing waits for an extra
click; a slide that sits still until you click again is the `delay="indefinite"` bug.
The 2026-09-27 run used the author's own animated 7-slide
deck, which is not published.

### L2 — leave a slide show running

Start the show again and leave it on screen. The bridge attaches to the running show
through LibreOffice's automation interface.

---
## Phase A — does COM automation work at all? (2 min)  *[needs PowerPoint]*

This is the foundation check. If it fails, stop; the architecture needs rethinking
before anything else is worth building.

Close LibreOffice first: if PowerPoint automation fails, the bridge falls back to
driving Impress, which could look like a pass. If PowerPoint was just installed, open
it once and answer its first-run privacy notice; until then it rejects automation.

1. Double-click `Start Speech Rehearsal.cmd`. If `demo-deck.pptx` is already open in
   PowerPoint, the bridge uses it. If nothing is open, a file picker asks "Which
   presentation are you rehearsing?": choose `demo-deck.pptx`.

Look for these four things:

| Expected | Meaning |
|---|---|
| PowerPoint shows the deck | COM connection succeeded |
| The slide show starts full screen | `SlideShowSettings.Run()` worked |
| Console: `PowerPoint slide show running - animations live. (demo-deck.pptx)` | COM mode, not fallback |
| A browser opens the rehearsal page, reading **Live PowerPoint ✓** | Token handoff via URL fragment worked |

**If the console says `Could not drive PowerPoint`**, copy the error message. The
most likely cause is a Microsoft Store install of Office, which has restricted COM
support compared to the desktop installer. With LibreOffice closed and no `-Target`,
the bridge then has nothing to drive, and opens the app on its own.

2. With the rehearsal page focused, press Page Down on the clicker. The slide show
   should advance and animate. Press `b`; the show should go black. Press `b` again.

Phase A passing is the real go/no-go for this project.

---

## Phase B — the vulnerability, in v1 (3 min)

1. `Ctrl+C` in the bridge console to stop v2.
2. Start the old version:

```powershell
# with LibreOffice today:
powershell -ExecutionPolicy Bypass -File .\rehearsal-bridge.v1.ps1 -Target Impress

# with PowerPoint tomorrow, drop the -Target and it uses COM:
powershell -ExecutionPolicy Bypass -File .\rehearsal-bridge.v1.ps1
```

**In keystroke mode `/goto` always sends Page Down**, so the show advances by one
slide rather than jumping to the number requested. Movement is still the proof —
the request was accepted and acted on.

3. Open `poc.html` from disk. Leave the token box empty. Click **Run tests**.

| Test | Expected against v1 |
|---|---|
| 1 — image tag GET | **The slide show moves** (jumps to slide 1 under COM; advances one under keystrokes) |
| 2 — unauthenticated `/health` | **READABLE**, shows `mode: com` |
| 3 — no token | accepted, slide moves |
| 4 — wrong token | accepted, slide moves |

Screenshot the PoC page next to the moving slide show. That is the evidence for the
writeup, and test 1 is the one that lands: no JavaScript, no CORS, just an image tag.

---

## Phase C — the fix, in v2 (3 min)

1. `Ctrl+C`. Start the current bridge again with the `.cmd`.
2. Copy the session token from the console.
3. Open `poc.html`, paste the token into the box, click **Run tests**.

| Test | Expected against v2 |
|---|---|
| 1 — image tag GET | no movement at all (405) |
| 2 — unauthenticated `/health` | blocked, or 401 |
| 3 — no token | blocked before sending, or rejected |
| 4 — wrong token | **401** |
| 5 — real token | **200, slide show jumps to slide 3** |

### Test 5 is not optional

If tests 1–4 fail and test 5 also fails, you have **not** proved the bridge is
secure — you have proved something is blocking every request, and it may be the
browser rather than the bridge. Recent Chrome versions have been tightening what a
page may request from local addresses.

Test 5 passing while 1–4 fail is the only result that means the access controls are
doing the work. If test 5 fails too, try Firefox and compare.

---

## Phase D — host and origin validation (2 min)

JavaScript cannot set the `Host` header, and a page opened from disk cannot fake its
`Origin`, so these need a terminal.

### D1 — a bad Host is refused

**In PowerShell, use `curl.exe`, not `curl`** — `curl` is an alias for
`Invoke-WebRequest`, which takes different arguments and will confuse you.

```powershell
curl.exe -i -X POST "http://127.0.0.1:8765/goto?n=2" -H "Host: evil.example" -H "X-Rehearsal-Token: PASTE_REAL_TOKEN"
```

Expected: `HTTP/1.1 403 Forbidden` and `{"ok":false,"error":"bad host"}`, with the
console logging `! rejected host: evil.example`.

This is the DNS-rebinding defence, and note what it demonstrates: **a valid token was
not enough.** Both checks are load-bearing.

Sanity check the other direction — same command with `Host: 127.0.0.1:8765` should
return 200 and move the slide.

### D2 — a hostile site's preflight is refused

`poc.html` is opened from disk, so its requests carry `Origin: null`, which the bridge
allows. Phase C therefore never sees a preflight refused. This sends the preflight a
page on some other website would send before it could attach the token header:

```powershell
curl.exe -i -X OPTIONS "http://127.0.0.1:8765/goto?n=2" -H "Origin: https://evil.example" -H "Access-Control-Request-Method: POST" -H "Access-Control-Request-Headers: X-Rehearsal-Token"
```

Expected: `HTTP/1.1 403 Forbidden` and `{"ok":false,"error":"origin not allowed"}`,
**no** `Access-Control-Allow-Origin` header in the reply, the console logging
`! rejected origin: https://evil.example`, and the slide not moving. A browser that
gets this answer never sends the real request, so the hostile page cannot use the
custom header at all.

Sanity check the other direction — same command with `-H "Origin: null"` should
return `204 No Content` with `Access-Control-Allow-Origin: null` and
`Access-Control-Allow-Headers: X-Rehearsal-Token`. A preflight never moves the slide,
either way.

---

## Phase E — idle shutdown (optional, 30 min unattended)

Start the bridge with a short timeout and walk away:

```powershell
powershell -ExecutionPolicy Bypass -File .\rehearsal-bridge.ps1 -IdleMinutes 2
```

After two minutes with no commands it should print `Idle for 2 minutes - shutting
down.` and exit.

---


---

## Optional — watch the keystroke-mode finding happen

`SECURITY.md` claims that in keystroke mode, a failed window match sends the
keystroke to whatever window has focus. You can watch it.

Run v1 with a target that matches nothing:

```powershell
powershell -ExecutionPolicy Bypass -File .\rehearsal-bridge.v1.ps1 -Target NoSuchWindow
```

Open a text editor, click into it, then trigger `poc.html` test 1 from another
window. Watch where the Page Down lands. That is the escalation, demonstrated
rather than asserted — and it is why COM mode is the better path.

## Record the results

**Run of 2026-09-27**, Windows 10, LibreOffice Impress, Edge. L1 and the clicker test
used the author's own animated 7-slide deck, which is not published. **Phase A** was
run later that night (2026-09-28, about 01:10), once PowerPoint was installed, with
`demo-deck.pptx`. The first pass found four bugs (listed under the table). After the fixes,
Phases C and D were run again against the final bridge, and those are the results
shown.

| Phase | Result | Notes |
|---|---|---|
| Antivirus — reaction to the bridge (product, detection, time; or none) | none seen | no Bitdefender alert for v1 or v2 during the run; v2 has a path exception from 2026-09-26 |
| L1 — animations play in Impress | PASS | all 7 slides matched the table |
| L1 — slides 1 and 6: typing and wipes play | PASS | |
| L1 — no slide waits for a click to build | PASS | one click per slide |
| A — COM automation | **PASS on the second attempt** | Run 2026-09-28 ~01:10, PowerPoint (desktop Microsoft 365) with `demo-deck.pptx`. First attempt FAILED: the bridge set `$ppt.Visible = $true`, which PowerShell can't convert to Office's `MsoTriState`, so it abandoned PowerPoint (bug 5 below). After the fix, the console printed "Opening demo-deck.pptx ..." and "PowerPoint slide show running - animations live.", `/health` answered mode `com`, and the show ran full screen on the second monitor |
| A — clicker advances slides | **PASS with the physical clicker** (after bug 6) | ~01:26: with the page in Edge showing **Live PowerPoint ✓**, the tester's physical clicker moved the script and the PowerPoint show together, forward and back across all 4 slides. The bridge log recorded every step, and Presenter View's "slide 3 of 4" matched the page's "SLIDE 3". The page happened to hold a 7-slide speech rather than the 4-slide sample: past slide 4 the script went on while the show held its last slide. The first attempt at this is what exposed bug 6. Earlier, the page's exact requests (`POST /goto`, `/blank`, `/unblank`) moved PowerPoint 1→2→3→2, blanked it (state 3) and restored it (state 1), each read back from PowerPoint. Blank *from the clicker* wasn't reported. Phases C and D repeated in PowerPoint mode passed |
| Impress — clicker advances script and show | PASS | physical clicker ("Page Down"): app and Impress both reached slide 7 of 7 together; Back matched too. Small visible delay per slide |
| B — v1 exploitable via `<img>` | PASS (vulnerable, as intended) | v1 console logged tests 1, 3 and 4 as accepted with no valid token |
| B — v1 `/health` readable | PASS (vulnerable, as intended) | 200 with `Access-Control-Allow-Origin: *` to Origin `https://evil.example` |
| C — v2 rejects 1–4 | PASS | 405, 405, 401, 401; the show didn't move (slide read back from Impress each time) |
| C — v2 accepts test 5 | PASS | 200, the show jumped to slide 3 |
| D — bad Host rejected 403 | PASS | with the real token; show didn't move |
| D — good Host accepted 200 | PASS | show moved |
| D — hostile-origin preflight refused 403 | PASS | `{"error":"origin not allowed"}`, no `Access-Control-Allow-Origin` |
| D — `null`-origin preflight answered 204 | PASS | `Access-Control-Allow-Origin: null`, `Allow-Headers: X-Rehearsal-Token` |
| E — idle shutdown | PASS | `-IdleMinutes 2`: "Idle for 2 minutes - shutting down." two minutes after start |
| App — split remembered after closing Edge | PASS | "Continue with…" and a hand-moved boundary both survived a full Edge restart |

Bugs the first pass found, all fixed the same evening:
1. **Keystroke mode can't drive an Impress show.** `-Target Impress` matched the editing
   window, and the show window (`SALTMPSUBFRAME`) can't take focus at all, so the keys
   never arrived. Fix: the bridge now drives Impress through LibreOffice's automation
   interface ("uno" mode), with no keystrokes; `/goto` became a jump, so Back works too.
2. **Port fallback.** With 8765 busy, the bridge listened on port 1: `@($Port, $Port + 1,
   $Port + 2)` is 8765, 8765, 1, 8765, 2 in PowerShell. Fixed with parentheses.
3. **Token handoff.** Windows dropped the `#t=…` fragment when opening the page, so it
   arrived without the token. The bridge now hands the address to the default browser's
   own program. Re-run: the new tab connected by itself.
4. **"Token rejected — click to re-enter" didn't re-ask.** Fixed in the page.
5. **PowerPoint mode never started (Phase A, first attempt).** `$ppt.Visible = $true`
   throws in PowerShell ("Cannot convert value True to type MsoTriState"), and it sat
   inside the try block, so the bridge gave up on PowerPoint. Fixed: `Visible = -1`
   (`msoTrue`), and the bridge now waits up to 10 s for the show window instead of 0.9 s.
6. **A slide number past the end of the deck killed the connection (PowerPoint mode).**
   With more script parts than slides, the page asked for slide 5 of a 4-slide deck.
   `GotoSlide(5)` threw, the bridge sent no reply, and the page showed "Bridge lost".
   Fixed: the number is capped at the last slide (as Impress mode already did), and
   PowerPoint errors now produce an `ok:false` reply instead of silence. Re-checked:
   requests for slides 5 and 7 answered 200 and the show stayed on slide 4.

**Start-up re-test, 2026-09-28 about 02:15.** This followed the change to how the
bridge finds its slide show: attach to the open deck, else a file picker, else the
app alone. All three paths were checked, with the bridge's slide and security checks
repeated:
- **Deck already open in PowerPoint:** used, with no "Opening…". The show started.
  Slides 3, too-far 9 (held on 4), blank and unblank all read back from PowerPoint.
  Image-tag GET 405, no token 401, wrong token 401, bad Host 403, hostile preflight
  403, `null` preflight 204, and the show didn't move for any of them.
- **`-Deck tools/demo-deck.fodp` with LibreOffice closed:** the bridge started
  LibreOffice, opened the file in Impress and started the show (mode `uno`). Slides 3
  and 2 read back from Impress; no token 401.
- **Nothing open, via `Start Speech Rehearsal.cmd`:** the picker appeared, the tester
  chose `demo-deck.pptx`, and PowerPoint opened it and ran the show. The new tab
  connected on its own (bridge log `-> slide 1`) and listed the saved rehearsals.

Also seen in Phase A: **PowerPoint's first-run privacy notice blocks automation.**
Until it's answered, PowerPoint rejects calls (`RPC_E_CALL_REJECTED`) and ignores
input to the show. Answer it before testing. It didn't cause bug 5: that error
reproduces with the notice closed.

Also seen: v1 sent three Page Downs close together and the show moved two, since
keystrokes are lossy. F5 didn't start Impress's show on this keyboard (a function-key
layer); the Slide Show menu did.

Anything that fails: copy the exact console line. The error text is worth more than
a description of it.
