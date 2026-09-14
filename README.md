# Portafolio RPA — AutoHotkey v2

Robots de automatización de escritorio (RPA) y API en AutoHotkey v2.

**Autor:** [Tu Nombre] — Senior QA Automation Engineer | RPA Developer

**Stack:** AutoHotkey v2 · UI Automation (UIA) · COM (Excel/Outlook) · CDO/SMTP · REST APIs (MSXML2.XMLHTTP) · Control de red por CLI

## Proyectos

| # | Carpeta | Qué hace | Técnica |
|---|---------|----------|---------|
| 01 | `01-uia-calculadora/` | Clics invisibles a la Calculadora de Windows (7+3) sin mover el mouse | UIA (ElementFromHandle, FindFirst, Click) |
| 02 | `02-com-excel/` | Crear libro de Excel, escribir celdas, fórmulas y formato desde el objeto COM (incluye variante `GetActiveObject` para un Excel ya abierto) | COM (Excel.Application) |
| 03 | `03-com-outlook-invisible/` | Enviar correo con adjunto usando Outlook en segundo plano (100% invisible) | COM (Outlook.Application) |
| 04 | `04-smtp-sin-outlook/` | Enviar correo directo por SMTP (CDO.Message) sin tener Outlook instalado | CDO + SMTP SSL |
| 05 | `05-rest-api/` | Consumir una API REST por GET y validar la respuesta (status 200) | MSXML2.XMLHTTP.6.0 |
| 06 | `06-tmac-mac-cli/` | Cambiar la dirección MAC del adaptador por línea de comandos (sin UI) | CLI backend (RunWait) |

## Cómo ejecutar

1. Requiere [AutoHotkey v2](https://www.autohotkey.com/) instalado.
2. Cada proyecto es un script `.ahk` independiente. Haz doble clic o ejecuta desde terminal:
   ```bash
   autohotkey.exe "01-uia-calculadora\Calculadora_UIA.ahk"
   ```
3. Los scripts que dependen de `UIA.ahk` (librería de [Descolada](https://github.com/Descolada/UIA-v2), MIT) la llevan junto al script en la misma carpeta.

## Configuración requerida

| Proyecto | Variable a rellenar |
|----------|---------------------|
| `03-com-outlook-invisible` | `oMail.To` → tu correo destino |
| `04-smtp-sin-outlook` | `CorreoDesde`, `CorreoPara`, `ContrasenaGmail` → tu correo y tu **contraseña de aplicación** de Gmail |
| `06-tmac-mac-cli` | `Adaptador` → nombre de tu conexión de red (`Get-NetAdapter` en PowerShell) |

> **Seguridad:** nunca subas contraseñas o datos reales a un repositorio público. Usa variables de entorno o placeholders.

## Notas técnicas

- **Decisión de técnica por app:** UIA si la app expone árbol de accesibilidad; COM si expone objeto; CLI/backend si tiene línea de comandos; `SendMessage` como última opción.
- **Uso de `ComObject("MSXML2.XMLHTTP.6.0")`** para peticiones HTTP sincrónicas en silencio.
- **`RunWait`** para ejecutar CLIs en segundo plano y capturar el código de salida.

## Licencia

Uso libre con fines de demostración y educativo. La librería `UIA.ahk` es MIT © Descolada.