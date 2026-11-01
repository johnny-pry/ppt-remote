# PPT Remoto

Seguí y controlá una presentación de PowerPoint desde cualquier tablet, celular o laptop de la red, con una página web.
Pensado para que quien presenta vea lo que viene y pase las diapositivas desde donde está, mientras el operador
sigue en la computadora conectada al proyector.

Funciona en Windows con Microsoft PowerPoint instalado.

## Instalación

Descargá `PPTRemoto.exe` de la última versión en [Releases](../../releases/latest) y copialo a la laptop del PowerPoint
(por ejemplo al Escritorio). No hace falta instalar nada más.

Como el programa no está firmado, Windows puede mostrar “Windows protegió su PC”: tocá **Más información** → **Ejecutar de todas formas**.

La primera vez que lo abrís, Windows pide permiso de administrador: decí **Sí**. Así el programa crea la regla del firewall
(“PPT Remoto”) que deja entrar a los otros dispositivos. No vuelve a preguntar, salvo que muevas el .exe a otra carpeta.

## Uso

1. Abrí la presentación en PowerPoint.
2. Abrí `PPTRemoto.exe`. No aparece ninguna ventana negra: se abre el **panel del operador** en el navegador
   y queda un ícono azul junto al reloj de Windows (clic derecho → abrir el panel, ver el PIN o **Salir**).
3. Quien presenta escanea el QR del panel con la tablet o el celular (ya incluye el PIN). Desde una laptop escribe la dirección y el PIN.
4. El operador inicia la presentación como siempre (o con el botón “Iniciar presentación”).

Quien presenta ve lo que está en pantalla y lo que viene, **incluidas las animaciones** (por ejemplo, cada viñeta que aparece),
sus notas, un reloj y un cronómetro. Puede pasar con los botones, deslizando el dedo, con el teclado o con un control remoto
(clicker), saltar a cualquier diapositiva, poner la pantalla en negro y ver la página en pantalla completa (botón abajo a la derecha).

El operador puede **desactivar el control remoto** en cualquier momento; quien presenta sigue viendo.

## Mensajes

El operador puede escribirle a quien presenta desde el panel (o usar los mensajes rápidos). En la tablet aparece un aviso grande
con respuestas rápidas (👍 Ok, Sí, No, Un momento) o para escribir. El operador ve “visto ✓” cuando el mensaje fue leído.
No hay sonidos a propósito: el audio de la computadora suele ir a los parlantes de la sala.
Los mensajes se borran al cerrar PPT Remoto.

## Bueno saber

- Los dispositivos y la computadora del PowerPoint tienen que estar en la misma red.
- En la tablet o el celular, alargá el tiempo de apagado de pantalla.
- Las vistas de animación se preparan en segundo plano unos segundos después de abrir la presentación.
  Se muestran las animaciones de entrada y salida; las de énfasis o movimiento no cambian la vista previa.
- Si editás diapositivas con la presentación abierta, tocá “Actualizar vistas previas” en el panel.
- Registro de errores: `%LOCALAPPDATA%\PPTRemoto\registro.txt`.

## Opciones (variables de entorno)

- `PPT_REMOTO_PIN=1234` fija el PIN (si no, cambia cada vez que se abre el programa).
- `PPT_REMOTO_PUERTO=8000` cambia el puerto.

## Desarrollo

```
pip install -r requirements.txt
python ppt_remote.py
construir.bat        → dist\PPTRemoto.exe
```

Cada push a GitHub construye `PPTRemoto.exe` en Actions. Para publicar una versión nueva en Releases:

```
git tag v1.0.0
git push origin v1.0.0
```

## Licencia

[MIT](LICENSE)
