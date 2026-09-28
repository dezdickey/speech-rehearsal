#!/usr/bin/env python3
"""Parse a marked-up speech script into the data the rehearsal app runs on.

This is the reference implementation for roadmap item 1 (text import). Porting
it to JavaScript inside speech-rehearsal.html is what turns the app from
one person's speech tool into a tool. No model is required: the slide splits
are user-supplied markers and the timing is arithmetic.

Markup the script uses:
    [SLIDE n]      slide boundary
    //             beat, about 1 second
    ///            full stop, about 3 seconds
    [[ ... ]]      stage direction (what the speaker's hands do)
    **bold**       stress

Timing: each slide's target is its share of the total word count, scaled to
TOTAL seconds. Stage directions are excluded from the count because they are
not spoken.
"""
import re, json, html, sys

TOTAL = 300          # seconds for the whole speech, pauses included
WPM   = 140          # measured speaking pace

BEAT = '<i class="beat" title="beat, about 1 second"></i>'
STOP = '<i class="stop" title="full stop, about 3 seconds"></i>'


def split_by_slide(body):
    marks = [(m.start(), int(m.group(1)))
             for m in re.finditer(r'`\[SLIDE (\d+)[^\]]*\]`', body)]
    out = []
    for k, (pos, n) in enumerate(marks):
        end = marks[k + 1][0] if k + 1 < len(marks) else len(body)
        out.append((n, body[pos:end]))
    return out


def clean(s):
    s = re.sub(r'`\[SLIDE[^\]]*\]`', '', s)
    s = re.sub(r'^\s*###.*$', '', s, flags=re.M)
    s = re.sub(r'^\s*>\s*\*\*Transition:\*\*', '', s, flags=re.M)
    s = re.sub(r'^\s*>\s*', '', s, flags=re.M)
    s = re.sub(r'^\s*(---|##).*$', '', s, flags=re.M)
    return re.sub(r'\n{3,}', '\n\n', s).strip()


def spoken_words(s):
    """Stage directions are not spoken, so they don't count toward timing."""
    return len(re.sub(r'\[\[.*?\]\]|[*`/]', ' ', s).split())


def direction(m):
    txt = m.group(1)
    room = txt.lower().startswith('turn back')
    return '<em class="do %s">%s %s</em>' % (
        'room' if room else 'board', '\u25cf' if room else '\u25b6', txt)


def to_html(s):
    s = html.escape(s)
    s = re.sub(r'\*\*(.+?)\*\*', lambda m: '<strong>%s</strong>' % m.group(1), s)
    s = re.sub(r'\[\[(.+?)\]\]', direction, s)
    s = s.replace('///', STOP).replace('//', BEAT)
    return ''.join('<p>%s</p>' % p.strip().replace('\n', ' ')
                   for p in s.split('\n\n') if p.strip())


def parse(text, start_marker=None, end_marker=None):
    """Optionally keep only the part between two marker lines (e.g. an outline's
    first spoken section and its reference list)."""
    a = text.index(start_marker) if start_marker and start_marker in text else 0
    b = text.index(end_marker) if end_marker and end_marker in text else len(text)
    body = text[a:b]
    chunks = [(n, clean(c)) for n, c in split_by_slide(body)]
    counts = {n: spoken_words(c) for n, c in chunks}
    total = sum(counts.values()) or 1
    return [
        dict(n=n,
             target=round(counts[n] / total * TOTAL),
             words=counts[n],
             full=to_html(c))
        for n, c in chunks
    ]


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else 'speech.md'
    steps = parse(open(src, encoding='utf-8').read())
    words = sum(s['words'] for s in steps)
    print('%d slides | %d words | %.2f min at %d wpm'
          % (len(steps), words, words / WPM, WPM), file=sys.stderr)
    print(json.dumps(steps, indent=2))
