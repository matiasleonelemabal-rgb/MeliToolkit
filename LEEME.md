# LA TRIBU · ML Toolkit — Tema de Streamlit

Este paquete lleva tu app a la estética **industrial oscura, amarillo + negro** (dirección 1a).

## Archivos

```
.streamlit/config.toml   → tema base (colores, modo oscuro)
estilos.py               → fuentes + CSS de marca + helpers
app.py                   → TU app YA modificada (lista para reemplazar)
```

> **Atajo:** si no querés tocar nada a mano, reemplazá tu `app.py` por el de este
> paquete y copiá `.streamlit/config.toml` + `estilos.py`. Ya tiene el tema,
> el logo del sidebar y los encabezados con barra amarilla integrados.
> Igual dejo abajo los pasos manuales por si preferís hacerlo vos.

## Cómo aplicarlo (3 pasos)

**1. Copiá los dos archivos al repo**, respetando las rutas:
- `.streamlit/config.toml` en la raíz del proyecto (misma carpeta que `app.py`).
- `estilos.py` en la raíz del proyecto.

**2. En `app.py`, importá y aplicá el tema** justo después de `set_page_config`:

```python
import streamlit as st
from estilos import aplicar_tema, encabezado, logo_sidebar

st.set_page_config(page_title="ML Toolkit", page_icon="📦", layout="wide")
aplicar_tema()          # <- pinta toda la app con el estilo de marca
```

**3. (Opcional pero recomendado) Usá los helpers de marca:**

- Reemplazá `st.sidebar.title("ML Toolkit")` por el logo real:
  ```python
  logo_sidebar()
  herramienta = st.sidebar.selectbox("Elegí herramienta", [...])
  ```

- En cada herramienta, cambiá `st.title(...) + st.caption(...)` por el encabezado con barra amarilla:
  ```python
  # antes:
  # st.title("Actualizar Integraly")
  # st.caption("Actualiza precio, stock y estado por SKU.")
  # después:
  encabezado("Actualizar Integraly", "Actualiza precio, stock y estado por SKU.")
  ```

Eso es todo: botones amarillos, métricas como cards, uploaders y tablas oscuras, tipografía Archivo/Manrope.

## Nota sobre límites de Streamlit
El CSS apunta a las clases de versiones recientes de Streamlit (data-testid). Si actualizás Streamlit y algo se ve distinto, avisame y ajusto los selectores. La maqueta `ML Toolkit.dc.html` es la referencia visual de cómo debería quedar cada pantalla.
