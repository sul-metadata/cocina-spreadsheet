# Why the hand-over gives a path, not a link

Measured in the Claude Code desktop app on 2026-09-22 by building real files in each
location and clicking every form. Re-test rather than assume if the client changes —
the failure mode is a link that looks fine and does nothing.

| link form | clickable? | resolves? |
|---|---|---|
| Relative, `./`-prefixed, or absolute POSIX **inside** the working directory | yes | **yes** |
| Absolute POSIX or `file:///home/...` **outside** it | yes | no |
| `file:///C:/...` | yes | no |
| Absolute Windows (`C:\...` or `C:/...`) | **no** | — |
| UNC, with or without `file://` | **no** | — |
| File-attachment tool, file outside the working directory | yes | no |

Three conclusions:

- **Reachability is scoped to the working directory tree**, whatever the mechanism. A
  workbook in `Documents` or `Desktop` — the normal save locations — cannot be handed
  over as a working link.
- **`.xlsx` never previews** even when the path resolves, because the app cannot render
  a spreadsheet. A successful click only locates the file.
- **The attachment tool reports success for a file it cannot reach.** "Delivered" is
  not evidence anything arrived.

The cause is that the app resolves paths against the session root, which on this
machine is inside WSL, so a Windows path has nothing to resolve against.

Hence Step 6: the absolute path in a code span, so it is byte-accurate and selectable
for pasting into Explorer. The README tells the user what to do if a click opens the
embedded viewer instead.
