# PPT Remoto

**English** | [Español](README.es.md)

Follow and control a PowerPoint presentation from any tablet, phone or laptop on your network, through a web page.
Built so the presenter can see what's coming and change slides from wherever they are, while an operator stays
at the computer connected to the projector.

Runs on Windows with Microsoft PowerPoint installed. **The interface is in Spanish.**

## Installation

Download [`PPTRemoto.exe`](../../releases/latest/download/PPTRemoto.exe) (latest version, see [Releases](../../releases/latest))
and copy it to the PowerPoint computer (for example to the Desktop). Nothing else to install.

The program isn't signed, so Windows may show “Windows protected your PC”: click **More info** → **Run anyway**.

The first time you open it, Windows asks for administrator permission: say **Yes**. This lets the program create a firewall
rule (“PPT Remoto”) so other devices can connect. It won't ask again unless you move the .exe to another folder.

## Usage

1. Open the presentation in PowerPoint.
2. Open `PPTRemoto.exe`. No console window appears: the **operator panel** opens in the browser,
   and a blue icon sits next to the Windows clock (right-click → open the panel, see the PIN, or **Salir** to quit).
3. The presenter scans the QR code in the panel with a tablet or phone (the PIN is included). From a laptop, type the address and PIN.
4. The operator starts the slideshow as usual (or with the “Iniciar presentación” button).

The presenter sees the current and next slide, **including animations** (for example, each bullet as it appears),
their notes, a clock and a stopwatch. They can change slides with the buttons, by swiping, with the keyboard or a remote
clicker, jump to any slide, black out the screen, and go full screen (button at the bottom right).

The operator can **turn off remote control** at any time; the presenter can still see everything.

## Messages

The operator can write to the presenter from the panel (or use quick messages). The tablet shows a large alert
with quick replies (👍 Ok, Sí, No, Un momento) or a box to type an answer. The operator sees “visto ✓” once the message has been read.
There are no sounds on purpose: the computer's audio usually goes to the room speakers.
Messages are cleared when PPT Remoto closes.

## Good to know

- The devices and the PowerPoint computer must be on the same network.
- On the tablet or phone, increase the screen timeout.
- Animation previews are prepared in the background a few seconds after the presentation opens.
  Entrance and exit animations are shown; emphasis and motion-path animations don't change the preview.
- If you edit slides while the presentation is open, click “Actualizar vistas previas” (refresh previews) in the panel.
- Error log: `%LOCALAPPDATA%\PPTRemoto\registro.txt`.

## Options (environment variables)

- `PPT_REMOTO_PIN=1234` sets a fixed PIN (otherwise it changes every time the program starts).
- `PPT_REMOTO_PUERTO=8000` changes the port.

## Development

```
pip install -r requirements.txt
python ppt_remote.py
construir.bat        → dist\PPTRemoto.exe
```

Every push to GitHub builds `PPTRemoto.exe` in Actions. To publish a new version to Releases:

```
git tag v1.0.0
git push origin v1.0.0
```

## License

[MIT](LICENSE)
