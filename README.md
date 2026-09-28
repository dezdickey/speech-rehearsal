# Speech Rehearsal

A browser app for rehearsing a timed talk (your script, a split timer and a pace
bar), plus an optional local bridge that lets **one USB presenter clicker advance
both your script and a live PowerPoint or LibreOffice Impress slide show**.

It also contains a security flaw I shipped and then fixed. The writeup is in
[SECURITY.md](SECURITY.md), and both versions of the server are in the repo so
the before and after can be read side by side.

---

## The problem

Practising a timed talk with slides is awkward. A presenter clicker enumerates as
a USB keyboard, and keyboard input only reaches the window that has focus. So the
clicker can drive the slides, or it can drive a script/timer app, but not both.
You end up rehearsing the words in one pass and the slide timing in another, and
never rehearsing the thing you will actually do.

## The approach

The rehearsal page keeps focus and receives every click. When the slide changes it
sends a command to a small local HTTP server, which drives the slide show through
the presentation program's own automation interface.

```
  clicker ──▶ browser page ──HTTP──▶ bridge (PowerShell) ──automation──▶ PowerPoint or Impress
              (keeps focus)          127.0.0.1:8765                       (never focused)
```

Because the bridge issues instructions rather than simulating keystrokes, the slide
show never needs focus.

---

## What is here

| File | What it does |
|---|---|
| `speech-rehearsal.html` | The rehearsal app: setup, split editor, teleprompter, split timer, pace bar. Works on its own |
| `rehearsal-bridge.ps1` | The optional local server. Token auth, origin and host checks, drives PowerPoint or Impress |
| `Start Speech Rehearsal.cmd` | **Start here** when you have slides: starts the bridge, which connects to your presentation and opens the app |
| `demo-deck.pptx` | A 4-slide demo deck that pairs with the app's built-in sample speech |
| `tools/demo-deck.fodp` | The demo deck's editable source (LibreOffice flat XML) |
| `rehearsal-bridge.v1.ps1` | **Not for everyday use.** The original, deliberately vulnerable version, kept for the security writeup. Run it only for the test in `TESTING.md` |
| `poc.html` | Access-control test page; the security writeup's reproduction |
| `SECURITY.md` | Vulnerability writeup: what v1 got wrong and why it mattered |
| `TESTING.md` | The test plan, with the recorded results |
| `tools/parse-script.py` | The original Python parser for marked-up scripts; the app now does this itself |

---

## Start here

**With your slides (Windows, PowerPoint or LibreOffice Impress):** double-click
**`Start Speech Rehearsal.cmd`**. That's the whole start-up:

1. It uses the presentation you already have open, in PowerPoint or in Impress.
2. If none is open, it asks which one you're rehearsing (a normal Windows file
   picker), opens it, and starts the slide show. A `.pptx` opens in PowerPoint if you
   have it; an `.odp` opens in Impress.
3. It opens the app, already connected: the top bar reads **PowerPoint connected** or
   **Impress connected**, and all your saved rehearsals are there.

Keep its window open while you practise; it stops itself after 2 idle hours. If
there's no slide show to drive (you cancel the picker, or have neither program), the
app opens on its own anyway.

**Without slides** (or if your antivirus blocks the bridge): double-click
`speech-rehearsal.html`. The script, the clock and your clicker all work without the
bridge. Nothing leaves your computer, and there is no AI involved.

Behind the scenes, the bridge mints a session token and opens the page with it in
the URL fragment. The tab remembers it, so reloading stays connected. If you opened
the page yourself, **Connect slides** tells you whether the bridge is running. If it
isn't, it says to start the launcher; if it is, it asks for the token and says where
to find it: the "Speech rehearsal bridge" window, on the `Session token:` line.

Drag the slide show to your second screen, then click once on the rehearsal window
so the clicker is aimed at it. An amber bar appears across the bottom whenever that
window is not focused, because a clicker silently doing nothing is the failure mode
that ruins a run.

Your antivirus may object to the bridge; see [Antivirus](#antivirus) below.

## Your speech

Mark where each slide starts with `[SLIDE 1]`, `[SLIDE 2: Title]` …, or a line on
its own like `Slide 2 – Title`. Without markers, you say how many slides there are
and the speech is cut into equal parts for you to fix. Optional markup:

| Mark | Meaning |
|---|---|
| `//` | beat, about 1 second |
| `///` | full stop, about 3 seconds |
| `**word**` | stress (bold in a `.docx` counts too) |
| `[[point at the chart]]` | stage direction, not spoken |

The split then opens in an editor: the whole speech, each slide's words in its own
colour, and a labelled line where each slide starts. Drag a line and it snaps to
the nearest sentence start; click a word and the nearest line moves exactly
there. `Ctrl+Z` undoes. Guessed boundaries are dashed and flagged "check".

### Your rehearsals

Every speech you split is kept as its own rehearsal on this computer. It's never
written over another one. The home screen (**← Rehearsals** at the top left, or **M**)
lists them, with slide count and planned time. Leaving a run in progress asks first.

With two or more, you can **sort** them by **Last opened** (the default), **Last
edited**, **Newest first**, **Oldest first**, **Name (A–Z)** or **Longest first**, and
**find** one by name. Each row's date follows the sort ("opened", "edited" or "made").
Opening a rehearsal doesn't count as editing it, and renaming changes no dates. Your
chosen order is remembered. For each rehearsal:

- **Continue** or **Edit split** picks one up where you left it.
- Its **⋯** menu holds the rest:
  - **Rename** gives it a clearer name. New ones are named after the file, or the
    speech's first few words.
  - **Export** saves it as a `.rehearsal.json` file. **Open a file** opens that file
    on another computer, or after clearing your browser data.
  - **Delete** removes it.

Up to 30 are kept. **More → Save as .txt** also writes the speech with its slide
markers, and opening that file gives the same split back.

## Using it

The top bar holds the connection to your slides, the detail level, and a **More** menu:
**Edit text**, **Fix the split**, **Options**, **Save as .txt**, **Export rehearsal**
and **Reset run**. The app is light by default; Options → Theme has a dark look for dim
rooms. **All keys**, at the bottom right, lists every shortcut.

| Key | Action |
|---|---|
| `Page Down` / `Space` / `→` | Next slide — advances script and slide show, records a split |
| `Page Up` / `←` | Previous slide |
| `F5` | Start the clock |
| `b` / `.` | Blank the slide show |
| `R` | Reset the run |
| `1` `2` `3` | Detail level: full script, keywords, cue only |
| `E` | Edit this slide's text |
| `M` | Your rehearsals, or a new one |
| `,` | Options |

Logitech presenters send Page Down, Page Up, F5, Escape and `b` by default, so a
clicker works without configuration. The footer echoes every key it receives and
what it mapped to, which is how you find out what your particular unit sends.

Escape is deliberately inert mid-run. It is the clicker's stop button and easy to
hit by accident.

### Three detail levels

Full script, then keywords, then nothing but the slide title. Keywords is an
automatic outline: each sentence's first few words plus any stressed words. The
point is to wean off the text: when you can hold your target time on cue-only, you
have learned the talk rather than memorised a page.

### Timing

Per-slide targets come from each slide's word count at your speaking pace (default
140 words a minute), plus the marked pauses. The time limit, the pace, and the hold
time for a slide with no words are settings under Options, with an option to stretch
the targets to fill the limit. Advancing records a split; the run ends with a table
of target versus actual, and the previous two runs persist as ghost marks under the
pace bar so you can see whether you fixed a slide or just moved the problem elsewhere.

---

## LibreOffice Impress

The bridge drives Impress through LibreOffice's own automation interface (UNO, via
the COM bridge LibreOffice registers on Windows), the same way it drives PowerPoint:
no keystrokes, no focus changes, and Back jumps back exactly. It attaches to the show
that is already running, starts it if a presentation is open but not showing, or
opens the file you pick and starts that.

There is also a keystroke mode (`-Target <window title>`), kept as a last resort. It
cannot drive an Impress show, whose window never takes keyboard focus, and it carries
more risk; see the writeup.

---

## Antivirus

**Bitdefender flagged `rehearsal-bridge.ps1`** the first time it was run, as
`CMD:Heur.BZC.PZQ.Boxter.1196.BA2C9B7D`. That is a *heuristic* detection: the script
matched a pattern, not a known piece of malware. Other antivirus products may react
the same way.

**Why a heuristic would flag it.** The bridge does several things malware also does:
it is a PowerShell script that its launcher starts with the execution policy bypassed;
it opens a network listener; it drives other programs through their automation
interfaces; and its keystroke fallback sends keys to another window, using a small
piece of C# compiled at run time to switch window focus.

**What it actually does.** It listens on `127.0.0.1` only, so nothing off your machine
can reach it, and it obeys only requests carrying a random token minted for that
session (see [SECURITY.md](SECURITY.md)). It understands four commands: go to slide
*n*, blank, unblank, and a health check. At start it looks for a presentation open in
PowerPoint or Impress, or shows a file picker so you can choose one. It also reads
which browser is your default, so it can open the rehearsal page in it. It writes no files, sends nothing to the internet, and shuts itself down
after 2 idle hours. It is one PowerShell file of a few hundred lines;
read it before you run it.

**If your antivirus blocks it**, the app still works fully without it. Whether to
allow the bridge is your decision; an exception usually covers one exact file path,
so don't move or rename the file afterwards. `rehearsal-bridge.v1.ps1` is the
deliberately vulnerable version and is even more likely to be flagged: don't run it
except for the security test. Removing the patterns the bridge doesn't need (the
keystroke fallback and the policy bypass) is planned; see `PLAN.md`.

---

## Security

The bridge is an HTTP server on your machine, and browsers let any page you visit
talk to `127.0.0.1`. The current version requires a per-session token in a custom
header, validates `Origin` and `Host`, accepts only POST, binds to loopback, and
shuts down when idle.

The first version did none of that. [SECURITY.md](SECURITY.md) covers what was
wrong, a working proof of concept, why `localhost` is not a security boundary, and
what the fix does and does not solve.

## Testing

`TESTING.md` is the test plan, with results. It was run on 2026-09-27/28 on Windows 10
with Edge:
- **LibreOffice Impress:** the token, `Host` and preflight checks passed, v1's
  vulnerability reproduced as intended, and a physical clicker advanced the script and
  the Impress show together.
- **PowerPoint** (Phase A, with `demo-deck.pptx`): the bridge opened the deck and ran
  the show, and a physical clicker advanced the script and the PowerPoint show together,
  forward and back. The page's requests also blanked and restored it, read back from
  PowerPoint, and the same access checks passed. This took two fixes along the way.

The testing found six bugs, all fixed; `TESTING.md` lists them. The app has a
built-in self-test: call `selfTest()` in the browser console, or open the file with
`#selftest` on the end of its address.

## Limitations

- The bridge is Windows only: it depends on COM.
- If the script has more parts than the deck has slides, the show waits on its last
  slide while the script carries on. Matching them is up to you.
- The demo deck has no animations.

## Notes

Built with AI coding assistance. The problem definition, the design decisions,
the security review and all testing on real hardware are mine.

MIT licensed.
