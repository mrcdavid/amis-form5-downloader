how to install - visit cmd

1. pip install playwright pypdf
if not recognized, then 
1.1. python -m playwright install chromium

2. playwright install chromium

check if installed
3. pip show playwright

check if installed
4. pip show pypdf

check playwright version
5. playwright --version


Commands to run

to kill edge
1. taskkill /F /IM msedge.exe

run to open edge debugging mode - laptop code
2. "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\edge-amis"

run this to open edge debugging mode
2.1. "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222

if it doesn't work, run this
2.2. "C:\Program Files\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222

Or, if your Edge is installed under your user account:
2.3. "%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222

to verify
3. http://127.0.0.1:9222/json/version