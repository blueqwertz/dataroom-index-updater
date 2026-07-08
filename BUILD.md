# Build

PyInstaller kann **nicht** cross-kompilieren: Ein Windows-`.exe` wird auf Windows
gebaut, ein macOS-`.app` auf einem Mac. Der Spec (`IndexUpdater.spec`) erkennt das
Zielsystem selbst und erzeugt jeweils das passende Artefakt.

## macOS (`.app`)

Auf einem Mac:

```sh
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm IndexUpdater.spec
# Ergebnis: dist/IndexUpdater.app
```

Ohne Mac zur Hand: Der Workflow `.github/workflows/build-macos.yml` baut das
`.app` bei jedem Push nach `main`/`master` auf einem `macos-latest`-Runner und legt
es als Artefakt `IndexUpdater-macos` ab (Actions-Tab → Run → Artifacts). Manuell
über „Run workflow" (workflow_dispatch) auslösbar.

> Das Bundle ist **nicht signiert/notarisiert**. Beim ersten Start ggf.
> Rechtsklick → „Öffnen" bzw. Systemeinstellungen → Datenschutz & Sicherheit.

## Windows (`.exe`)

```sh
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm IndexUpdater.spec
# Ergebnis: dist/IndexUpdater.exe
```
