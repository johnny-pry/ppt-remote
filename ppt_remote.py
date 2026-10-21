"""
PPT Remoto — el predicador sigue y controla el PowerPoint desde su tablet o laptop.

Se ejecuta en la laptop del PowerPoint (Windows + Microsoft PowerPoint):
    python ppt_remote.py        (o PPTRemoto.exe)

- Página del predicador:  http://<ip-de-la-laptop>:8000/
- Panel del operador:     http://localhost:8000/operador   (solo desde la misma laptop)

Corre sin ventana: queda un ícono en la bandeja de Windows (junto al reloj) para abrir el panel o salir.
"""
import asyncio
import hashlib
import os
import queue
import re
import secrets
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path

import pystray
import pythoncom
import pywintypes
import segno
import uvicorn
import win32com.client
import win32con
import win32event
from PIL import Image, ImageDraw
from win32com.shell import shell, shellcon
from fastapi import FastAPI, HTTPException, Request, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, HTMLResponse, Response

PORT = int(os.environ.get("PPT_REMOTO_PUERTO", 8000))
PIN = os.environ.get("PPT_REMOTO_PIN") or f"{secrets.randbelow(10000):04d}"
CACHE = Path(tempfile.gettempdir()) / "ppt_remoto"
REGISTRO = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "PPTRemoto" / "registro.txt"
REGLA_FIREWALL = "PPT Remoto"
ANCHO_IMAGEN = 1280
INTERVALO = 0.25          # segundos entre lecturas del estado de PowerPoint
INTERVALO_DECK = 2.0      # segundos entre revisiones de cambios en la presentación

# Constantes de PowerPoint (PpSlideShowState / tipos de forma)
PP_CORRIENDO, PP_NEGRO, PP_FIN = 1, 3, 5
MSO_PLACEHOLDER, PP_PLACEHOLDER_CUERPO = 14, 2
PP_AL_HACER_CLIC, MSO_ANIM_SET, MSO_ANIM_VISIBILIDAD = 1, 8, 8
PP_GUARDAR_PPTX = 24


def base_dir() -> Path:
    # Con PyInstaller los archivos van dentro del .exe (sys._MEIPASS)
    return Path(getattr(sys, "_MEIPASS", Path(__file__).parent))


def ip_local() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except OSError:
        return socket.gethostbyname(socket.gethostname())


def conectar_powerpoint():
    """Devuelve la instancia de PowerPoint abierta, o None. Nunca abre PowerPoint por su cuenta."""
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application")
    except pythoncom.com_error:
        return None


def notas_de(diapo) -> str:
    try:
        formas = diapo.NotesPage.Shapes
        for i in range(1, formas.Count + 1):
            f = formas.Item(i)
            if f.Type == MSO_PLACEHOLDER and f.PlaceholderFormat.Type == PP_PLACEHOLDER_CUERPO:
                return f.TextFrame.TextRange.Text.replace("\r", "\n").strip()
    except pythoncom.com_error:
        pass
    return ""


def es_entrada(efecto) -> bool:
    """Las animaciones de entrada son las que vuelven visible la forma."""
    try:
        for j in range(1, efecto.Behaviors.Count + 1):
            b = efecto.Behaviors.Item(j)
            if b.Type == MSO_ANIM_SET and b.SetEffect.Property == MSO_ANIM_VISIBILIDAD and b.SetEffect.To == 1:
                return True
    except pythoncom.com_error:
        pass
    return False


def eventos_animacion(diapo):
    """Lee la secuencia principal: devuelve (cantidad de clics, [(clic, (forma, párrafo), visible)])."""
    secuencia = diapo.TimeLine.MainSequence
    indices = {diapo.Shapes.Item(i).Id: i for i in range(1, diapo.Shapes.Count + 1)}
    clic, eventos = 0, []
    for i in range(1, secuencia.Count + 1):
        efecto = secuencia.Item(i)
        if efecto.Timing.TriggerType == PP_AL_HACER_CLIC:
            clic += 1
        if efecto.Exit:
            visible = False
        elif es_entrada(efecto):
            visible = True
        else:
            continue   # énfasis o trayectoria: no cambian lo que se ve
        try:
            parrafo = efecto.Paragraph
        except pythoncom.com_error:
            parrafo = 0
        indice = indices.get(efecto.Shape.Id)
        if indice:
            eventos.append((clic, (indice, parrafo), visible))
    return clic, eventos


def exportar_paso(diapo, eventos, paso, ruta, alto):
    """Exporta la diapositiva tal como se ve después de `paso` clics."""
    visibles = {}
    for clic, objetivo, visible in eventos:
        visibles.setdefault(objetivo, not visible)   # antes de su primera animación
        if clic <= paso:
            visibles[objetivo] = visible
    copia = diapo.Duplicate().Item(1)
    try:
        for (indice, parrafo), visible in visibles.items():
            if visible:
                continue
            forma = copia.Shapes.Item(indice)
            if parrafo:
                texto = forma.TextFrame2.TextRange.GetParagraphs(parrafo, 1)
                texto.Font.Fill.Transparency = 1
                try:
                    texto.ParagraphFormat.Bullet.UseTextColor = True
                except pythoncom.com_error:
                    pass
            else:
                forma.Visible = False
        copia.Export(str(ruta), "PNG", ANCHO_IMAGEN, alto)
    finally:
        copia.Delete()


class PowerPoint(threading.Thread):
    """Único hilo que habla con PowerPoint (COM). Recibe órdenes por una cola y publica el estado."""

    def __init__(self, al_cambiar):
        super().__init__(daemon=True)
        self.ordenes: queue.Queue = queue.Queue()
        self.al_cambiar = al_cambiar
        self.estado = {"conectado": False, "deck": None, "presentando": False,
                       "diapo": None, "paso": 0, "pasos": 0, "negro": False, "fin": False}
        self.firma = None
        self.ultima_revision = 0.0
        self.generacion = 0   # cancela exportaciones viejas cuando cambia la presentación
        self.forzar_revision = False

    def run(self):
        pythoncom.CoInitialize()
        app = None
        while True:
            try:
                orden = self.ordenes.get(timeout=INTERVALO)
            except queue.Empty:
                orden = None
            try:
                app = app or conectar_powerpoint()
                if app is None:
                    self._publicar(conectado=False, deck=None, presentando=False, diapo=None)
                    self.firma = None
                    continue
                if orden:
                    self._ejecutar(app, orden)
                self._leer(app)
            except Exception:
                # PowerPoint cerrado, ocupado (diálogo abierto) o en transición: reintentar
                app = None

    # --- lectura de estado ---------------------------------------------------------------

    def _presentacion(self, app):
        if app.SlideShowWindows.Count:
            ventana = app.SlideShowWindows(1)
            return ventana.Presentation, ventana.View
        if app.Presentations.Count:
            try:
                return app.ActivePresentation, None
            except pythoncom.com_error:
                return app.Presentations(1), None
        return None, None

    def _leer(self, app):
        pres, vista = self._presentacion(app)
        forzar = self.forzar_revision
        if pres is None:
            self.firma = None
            self._publicar(conectado=True, deck=None, presentando=False, diapo=None)
            return
        if forzar or time.monotonic() - self.ultima_revision > INTERVALO_DECK \
                or pres.FullName != (self.firma or ("",))[0]:
            self.forzar_revision = False
            self._revisar_deck(pres, forzar)

        cambios = {"conectado": True, "presentando": vista is not None,
                   "diapo": None, "paso": 0, "pasos": 0, "negro": False, "fin": False}
        if vista is not None:
            estado = vista.State
            cambios["negro"] = estado == PP_NEGRO
            cambios["fin"] = estado == PP_FIN
            if not cambios["fin"]:
                try:
                    cambios["diapo"] = vista.Slide.SlideIndex
                    cambios["paso"] = vista.GetClickIndex()
                    cambios["pasos"] = vista.GetClickCount()
                except pythoncom.com_error:
                    pass
        self._publicar(**cambios)

    def _revisar_deck(self, pres, forzar=False):
        self.ultima_revision = time.monotonic()
        diapos = pres.Slides
        ids = tuple(diapos.Item(i).SlideID for i in range(1, diapos.Count + 1))
        firma = (pres.FullName, ids)
        if firma == self.firma and not forzar:
            return
        self.firma = firma
        self.generacion += 1
        deck_id = hashlib.sha1(f"{firma}{time.time()}".encode()).hexdigest()[:10]
        lista = []
        for i in range(1, diapos.Count + 1):
            d = diapos.Item(i)
            lista.append({"n": i, "oculta": bool(d.SlideShowTransition.Hidden), "notas": notas_de(d)})
        ancho, alto = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        self._publicar(deck={"id": deck_id, "nombre": pres.Name, "diapos": lista,
                             "listas": 0, "pasos": {}, "revisadas": 0, "proporcion": ancho / alto})
        threading.Thread(target=self._exportar, daemon=True,
                         args=(pres.FullName, deck_id, len(lista), round(ANCHO_IMAGEN * alto / ancho),
                               self.generacion)).start()

    def _exportar(self, nombre_completo, deck_id, total, alto, generacion):
        """Exporta cada diapositiva a PNG en su propio hilo, para no frenar el control."""
        pythoncom.CoInitialize()
        carpeta = CACHE / deck_id
        carpeta.mkdir(parents=True, exist_ok=True)
        copia = None
        try:
            app = conectar_powerpoint()
            pres = next(app.Presentations(i) for i in range(1, app.Presentations.Count + 1)
                        if app.Presentations(i).FullName == nombre_completo)
            # 1) La diapositiva completa: sirve de vista previa rápida y para la grilla
            for n in range(1, total + 1):
                if generacion != self.generacion:
                    return
                pres.Slides(n).Export(str(carpeta / f"{n}.png"), "PNG", ANCHO_IMAGEN, alto)
                self._actualizar_deck(deck_id, listas=n)

            # 2) Una imagen por cada clic de animación, en una copia invisible de la presentación
            ruta_copia = carpeta / "copia.pptx"
            pres.SaveCopyAs(str(ruta_copia), PP_GUARDAR_PPTX)
            copia = app.Presentations.Open(str(ruta_copia), ReadOnly=True, Untitled=False, WithWindow=False)
            # Empieza por la diapositiva que se está mostrando, para que lo inmediato esté listo primero
            inicio = self.estado["diapo"] or 1
            orden = list(range(inicio, total + 1)) + list(range(1, inicio))
            for revisadas, n in enumerate(orden, start=1):
                if generacion != self.generacion:
                    return
                diapo = copia.Slides(n)
                clics, eventos = eventos_animacion(diapo)
                for paso in range(clics + 1 if clics else 0):
                    exportar_paso(diapo, eventos, paso, carpeta / f"{n}_{paso}.png", alto)
                deck = self.estado["deck"]
                if deck and deck["id"] == deck_id:
                    pasos = {**deck["pasos"], str(n): clics} if clics else deck["pasos"]
                    self._actualizar_deck(deck_id, pasos=pasos, revisadas=revisadas)
        except Exception as e:
            print(f"No se pudieron exportar las diapositivas: {e}")
        finally:
            if copia is not None:
                try:
                    copia.Close()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def _actualizar_deck(self, deck_id, **cambios):
        deck = self.estado["deck"]
        if deck and deck["id"] == deck_id:
            self._publicar(deck={**deck, **cambios})

    def _publicar(self, **cambios):
        nuevo = {**self.estado, **cambios}
        if nuevo != self.estado:
            self.estado = nuevo
            self.al_cambiar(nuevo)

    # --- órdenes ---------------------------------------------------------------------------

    def _ejecutar(self, app, orden):
        tipo = orden.get("orden")
        pres, vista = self._presentacion(app)
        if tipo == "refrescar":
            self.forzar_revision = True
            return
        if tipo == "iniciar" and pres is not None and vista is None:
            pres.SlideShowSettings.Run()
            return
        if vista is None:
            return
        if tipo == "siguiente":
            vista.Next()
        elif tipo == "anterior":
            vista.Previous()
        elif tipo == "ir" and isinstance(orden.get("n"), int):
            if 1 <= orden["n"] <= pres.Slides.Count:
                vista.GotoSlide(orden["n"])
        elif tipo == "negro":
            vista.State = PP_CORRIENDO if vista.State == PP_NEGRO else PP_NEGRO


# --- servidor web ------------------------------------------------------------------------------

app = FastAPI()
clientes: dict[WebSocket, dict] = {}
config = {"control_remoto": True}
mensajes: list[dict] = []   # chat operador ↔ pastor (solo en memoria)
MAX_MENSAJES, MAX_LARGO = 50, 300
firewall_ok = True
bucle: asyncio.AbstractEventLoop | None = None


def instantanea() -> dict:
    return {**ppt.estado, **config, "conectados": sum(1 for c in clientes.values() if not c["operador"])}


async def difundir():
    datos = instantanea()
    for ws, info in list(clientes.items()):
        try:
            await ws.send_json({**datos, "puede_controlar": info["operador"] or (info["pin_ok"] and config["control_remoto"]),
                                "pin_ok": info["pin_ok"], "mensajes": mensajes if info["pin_ok"] else []})
        except Exception:
            clientes.pop(ws, None)


def al_cambiar(_estado):
    if bucle:
        bucle.call_soon_threadsafe(lambda: asyncio.ensure_future(difundir()))


ppt = PowerPoint(al_cambiar)


def es_local(host: str | None) -> bool:
    return host in ("127.0.0.1", "::1", "localhost")


def pagina(nombre: str) -> HTMLResponse:
    return HTMLResponse((base_dir() / "web" / nombre).read_text(encoding="utf-8-sig"))


@app.on_event("startup")
async def inicio():
    global bucle
    bucle = asyncio.get_running_loop()
    shutil.rmtree(CACHE, ignore_errors=True)
    ppt.start()


@app.get("/")
async def predicador():
    return pagina("predicador.html")


@app.get("/operador")
async def operador(request: Request):
    if not es_local(request.client.host):
        raise HTTPException(403, "El panel del operador solo se abre desde la laptop del PowerPoint.")
    return pagina("operador.html")


@app.get("/info")
async def info(request: Request):
    if not es_local(request.client.host):
        raise HTTPException(403)
    return {"pin": PIN, "url": f"http://{ip_local()}:{PORT}/", "url_con_pin": f"http://{ip_local()}:{PORT}/?pin={PIN}",
            "firewall": firewall_ok}


@app.get("/qr.svg")
async def qr(request: Request):
    if not es_local(request.client.host):
        raise HTTPException(403)
    svg = segno.make(f"http://{ip_local()}:{PORT}/?pin={PIN}").svg_inline(scale=8, border=2)
    return Response(svg, media_type="image/svg+xml")


@app.get("/diapos/{deck_id}/{nombre}.png")
async def diapo(deck_id: str, nombre: str):
    # nombre: "3" (diapositiva completa) o "3_2" (diapositiva 3 después de 2 clics)
    if not deck_id.isalnum() or not re.fullmatch(r"\d+(_\d+)?", nombre):
        raise HTTPException(404)
    ruta = CACHE / deck_id / f"{nombre}.png"
    if not ruta.is_file():
        raise HTTPException(404)
    return FileResponse(ruta, headers={"Cache-Control": "max-age=86400"})


@app.websocket("/ws")
async def ws(websocket: WebSocket):
    await websocket.accept()
    host = websocket.client.host if websocket.client else None
    es_operador = es_local(host) and websocket.query_params.get("operador") == "1"
    info = {"operador": es_operador, "pin_ok": es_operador or secrets.compare_digest(
        websocket.query_params.get("pin", "").encode(), PIN.encode())}
    clientes[websocket] = info
    await difundir()
    try:
        while True:
            msg = await websocket.receive_json()
            orden = msg.get("orden")
            if info["operador"] and orden == "control_remoto":
                config["control_remoto"] = bool(msg.get("valor"))
                await difundir()
            elif orden == "mensaje" and info["pin_ok"]:
                texto = str(msg.get("texto", "")).strip()[:MAX_LARGO]
                if texto:
                    mensajes.append({"id": (mensajes[-1]["id"] + 1) if mensajes else 1,
                                     "de": "operador" if info["operador"] else "pastor",
                                     "texto": texto, "hora": time.strftime("%H:%M"), "visto": False})
                    del mensajes[:-MAX_MENSAJES]
                    await difundir()
            elif orden == "visto" and info["pin_ok"]:
                # El que recibe marca como vistos los mensajes del otro lado
                de_quien = "pastor" if info["operador"] else "operador"
                for m in mensajes:
                    if m["de"] == de_quien:
                        m["visto"] = True
                await difundir()
            elif orden in ("siguiente", "anterior", "ir", "negro", "iniciar", "refrescar"):
                solo_operador = orden in ("iniciar", "refrescar")
                permitido = info["operador"] or (not solo_operador and info["pin_ok"] and config["control_remoto"])
                if permitido:
                    ppt.ordenes.put(msg)
    except (WebSocketDisconnect, RuntimeError):
        pass
    finally:
        clientes.pop(websocket, None)
        await difundir()


def puerto_ocupado() -> bool:
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", PORT)) == 0


def abrir_panel():
    webbrowser.open(f"http://localhost:{PORT}/operador")


def redirigir_salida():
    """Sin consola (el .exe) no hay dónde imprimir: los mensajes van a un archivo de registro."""
    if sys.stdout is None or sys.stderr is None:
        REGISTRO.parent.mkdir(parents=True, exist_ok=True)
        archivo = open(REGISTRO, "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stderr = archivo


def asegurar_firewall() -> bool:
    """Crea (una sola vez) la regla del firewall que deja entrar a las tablets. Pide permiso de administrador."""
    if os.environ.get("PPT_REMOTO_SIN_FIREWALL"):
        return True
    programa = sys.executable

    def regla_correcta():
        salida = subprocess.run(["netsh", "advfirewall", "firewall", "show", "rule", f"name={REGLA_FIREWALL}", "verbose"],
                                capture_output=True, encoding="oem", errors="replace",
                                creationflags=subprocess.CREATE_NO_WINDOW).stdout
        return programa.lower() in salida.lower()

    if regla_correcta():
        return True
    comando = (f'/c netsh advfirewall firewall delete rule name="{REGLA_FIREWALL}" & '
               f'netsh advfirewall firewall add rule name="{REGLA_FIREWALL}" dir=in action=allow '
               f'program="{programa}" enable=yes profile=any')
    try:
        proceso = shell.ShellExecuteEx(fMask=shellcon.SEE_MASK_NOCLOSEPROCESS, lpVerb="runas",
                                       lpFile="cmd.exe", lpParameters=comando, nShow=win32con.SW_HIDE)
        win32event.WaitForSingleObject(proceso["hProcess"], 60_000)
    except pywintypes.error:
        return False   # el usuario no dio permiso de administrador
    return regla_correcta()


def icono_bandeja(servidor: uvicorn.Server):
    imagen = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    dibujo = ImageDraw.Draw(imagen)
    dibujo.rounded_rectangle((2, 10, 62, 54), radius=8, fill=(79, 140, 255))
    dibujo.polygon([(26, 20), (26, 44), (45, 32)], fill="white")

    def salir(icono, _):
        servidor.should_exit = True
        icono.stop()

    menu = pystray.Menu(
        pystray.MenuItem("Abrir panel del operador", lambda *_: abrir_panel(), default=True),
        pystray.MenuItem(f"PIN: {PIN}", None, enabled=False),
        pystray.MenuItem("Salir", salir),
    )
    pystray.Icon("PPTRemoto", imagen, "PPT Remoto", menu).run()


def main():
    global firewall_ok
    if puerto_ocupado():   # ya está abierto: solo mostrar el panel
        abrir_panel()
        return
    redirigir_salida()
    firewall_ok = asegurar_firewall()
    print(f"{time.strftime('%Y-%m-%d %H:%M')}  PPT Remoto en http://{ip_local()}:{PORT}/  PIN {PIN}")
    servidor = uvicorn.Server(uvicorn.Config(app, host="0.0.0.0", port=PORT, log_level="warning", log_config=None))
    threading.Thread(target=servidor.run, daemon=True).start()
    threading.Timer(1.5, abrir_panel).start()
    icono_bandeja(servidor)
    os._exit(0)


if __name__ == "__main__":
    main()
