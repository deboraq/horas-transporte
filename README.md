# Horas Transporte

Sistema para el **delegado de transporte**: cargar asistencia diaria, calcular horas (normales, extras 50%/100%, descanso 100%, cenas) y consultar el mes por colaborador.

Período laboral: **del 26 al 25** de cada mes.

## Uso recomendado (Google Drive)

La app principal es `control_horas.html`, publicada como **aplicación web** de Google Apps Script. Los datos viven en una carpeta de Drive compartida.

### Configurar (una sola vez)

1. Entrá a [script.google.com](https://script.google.com) y creá un proyecto.
2. Pegá el contenido de `sheets/CodeDrive.gs` en `Código.gs`.
3. Creá un archivo HTML llamado `ControlHoras` y pegá el contenido de `control_horas.html`.
4. **Implementar → Nueva implementación → Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: según tu organización (p. ej. usuarios de Google Workspace)
5. Guardá la URL en favoritos.

Instrucciones detalladas: `COMO_USAR_DRIVE.txt`.

### Actualizar la app

Cada vez que cambies `control_horas.html`:

1. Reemplazá el contenido de `ControlHoras` en Apps Script.
2. **Implementar → Administrar implementaciones → lápiz → Nueva versión → Implementar**.

Si no creás una **nueva versión**, la URL sigue mostrando el código viejo.

### Qué hace la app

| Vista | Función |
|-------|---------|
| **Por día** | Cargar entrada, salida y citación; marcar feriado; franco |
| **Por colaborador** | Ver todos los días del período (vacíos incluidos); editar en la tabla; exportar Excel del colaborador |
| **Personal** | Alta/baja de legajos |

También: respaldo/restaurar JSON, Excel general, sync con Drive (`datos.json` + carpeta `Respaldos/`).

## Reglas de cálculo (resumen)

- **Lun–vie:** jornada 8 hs desde la citación · **Sábado:** 4 hs
- **Citación:** si llega antes, ese tiempo no cuenta
- **Normales:** dentro del horario, antes de las 21:00
- **Extras:** fuera del horario · diurnas al 50% · si alguna extra toca 21:00–06:00 → **todas** las extras al 100%
- **Domingo / feriado:** todo al 100% (mínimo 6 hs si trabajó)
- **Cena:** 1 si trabaja en franja nocturna (21:00–06:00)
- **Descanso 100%:** si hay menos de 12 hs entre jornadas, se paga la diferencia

## Uso local (sin Drive)

Doble clic en `ABRIR_CONTROL.bat` o abrí `control_horas.html` en el navegador. Los datos quedan en el navegador (`localStorage`); usá **Respaldo** / **Restaurar** para copiar entre PCs.

## Herramientas Python (opcional)

Para generar planillas Excel o comparar contra Legajos Online:

```bash
pip install -r requirements.txt
python src/generar_planilla.py --anio 2026 --mes 6
```

Ver también `COMO_USAR.txt` y los `.bat` de la carpeta.

## Estructura

```
Horas transporte/
  control_horas.html      ← app principal (copiar a Apps Script)
  ABRIR_CONTROL.bat
  COMO_USAR_DRIVE.txt
  sheets/
    CodeDrive.gs          ← backend Drive
    Code.gs               ← variante Sheets
  src/                    ← scripts Python
  data/                   ← plantillas / dumps locales
```

## Privacidad

No subas a GitHub archivos con datos reales de personal (`datos.json`, respaldos, Excel con legajos). Este repo es solo el programa.
