# SAP Invoice Approval Status Automation

Automatiza la validación diaria del estatus de aprobación de facturas en SAP, actualizando un Excel de seguimiento compartido y armando un borrador de correo de alerta cuando algo necesita atención.

> **Nota:** este es un repositorio de portafolio ya sanitizado. Todos los identificadores específicos de la empresa (rutas, correos, nodo de SAP) fueron reemplazados por variables de entorno o placeholders genéricos — ver `.env.example`. No se incluye ningún dato real de negocio en este repositorio.

## Problema

Un analista revisaba manualmente, documento por documento, si cada factura ya había sido autorizada en SAP, copiando el estatus a mano en un Excel de control y avisando por correo cuando algo se quedaba "atorado" sin resolver.

Este script automatiza todo el flujo de principio a fin.

## Qué hace

1. Abre el Excel real y lee la tabla `Facturas` (columnas `Comprobante` y `Fecha`).
2. Por cada comprobante, entra al reporte de aprobaciones en SAP (lista clásica de texto, no ALV Grid) y lee todas las filas de aprobación del documento — puede haber varias, una por cada aprobador en la cadena.
3. Valida que las filas leídas realmente correspondan al documento solicitado, como protección contra pantallas desincronizadas por popups inesperados.
4. Calcula un score:
   - `0` → Rechazado
   - `1` → Autorizado (estatus definitivo)
   - `0.5` → En proceso (puede tener aprobaciones intermedias, pero no la autorización final)
5. Escribe el score en la columna `Status` y los días transcurridos desde la fecha de la factura en `DiasTranscurridos`.
6. Si hay documentos rechazados, atorados (más de N días en proceso) o no encontrados, arma **un borrador** de correo con prioridad alta — nunca se envía solo, se abre para revisión manual antes de dar clic en Enviar.
7. Si todo está en orden, no genera ningún correo.

## Candado de seguridad

El correo se genera siempre como borrador (`mail.Display()`), nunca se envía automáticamente. El cambio a envío automático (`mail.Send()`) es una decisión manual que se toma directamente en el código, una vez que se ha confirmado que el reporte es confiable tras varias corridas.

## Requisitos

- Windows, con SAP GUI Scripting habilitado
- Una sesión de SAP ya abierta y logueada (el script no abre una nueva)
- Outlook instalado (para el borrador de correo)
- Python 3.9+

```bash
pip install -r requirements.txt
```

## Configuración

1. Copia `.env.example` a `.env`.
2. Llena tus valores reales: ruta del Excel, nombre de columnas, correo destino, nodo de SAP, etc.
3. El archivo `.env` nunca se sube al repo (está en `.gitignore`).

## Uso

1. Cierra el Excel si lo tienes abierto.
2. Abre SAP Logon, entra a tu sesión y déjala en la pantalla inicial.
3. Corre el script:

```bash
python invoice_approval_status_automation.py
```

## Estructura

```
.
├── invoice_approval_status_automation.py   # script principal
├── .env.example                            # plantilla de configuración
├── requirements.txt
└── .gitignore
```

## Notas técnicas

- La lectura de SAP usa `GuiLabel` en coordenadas fijas de fila/columna (lista clásica), no un grid ALV — por eso el filtrado se hace en la pantalla de selección, antes de ejecutar.
- El script es idempotente: si un documento ya quedó `Autorizado` en una corrida anterior, no se vuelve a consultar en SAP.
- Guarda el progreso cada N documentos (configurable) para no perder avance si algo falla a mitad de la corrida.
