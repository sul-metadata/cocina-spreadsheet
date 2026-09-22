# Installing Python

Read this only when the Step 0 probe finds no interpreter. The builder is standard
library only — no `pip install`, no virtualenv, no `openpyxl` — so Python itself is the
whole dependency, and any Python 3 will do.

**Don't install anything yourself.** It needs the user's consent and often their admin
rights. Offer the option that fits their machine:

- **Windows, official build** — download Python 3 from
  [python.org/downloads](https://www.python.org/downloads/) and tick *Add python.exe to
  PATH* during setup. Or from a terminal:
  ```
  winget install Python.Python.3.13
  ```
  If a real Python is installed but `python` still hits the Store stub, the alias is
  shadowing it: turn it off under *Settings → Apps → Advanced app settings → App
  execution aliases*.
- **Windows, borrowing WSL's interpreter** — if WSL is installed but the distro has no
  Python:
  ```
  wsl -d <distro> -- sudo apt install -y python3
  ```
  If WSL itself is missing, `wsl --install` sets it up but needs a restart, so the
  python.org installer is usually quicker.
- **macOS** — `brew install python`, or the python.org installer. `xcode-select
  --install` also provides a `python3`.
- **Linux** — `sudo apt install -y python3`, or `sudo dnf install -y python3`.

Once they confirm it's installed, re-run the probe to see which command works, then
carry on. If they'd rather not install anything, say plainly that the workbook can't be
built without an interpreter — there is no fallback that avoids it.

## The Store-stub trap

On Windows, `python` may be a 0-byte Microsoft Store alias stub at
`%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe`. The command exists, so
`Get-Command python` succeeds and an exit-status check passes, but running it prints
"Python was not found" and exits 9009. Always read the probe's output rather than its
status.
