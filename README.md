# 🚀 Fixy Dashboard — Guía de uso

## Estructura del proyecto

```
fixy_dashboard/
├── app.py                  ← Backend (servidor Flask)
├── requirements.txt        ← Dependencias Python
├── INICIAR.bat             ← Doble-click en Windows para arrancar
├── iniciar.sh              ← En Mac/Linux para arrancar
├── templates/
│   └── dashboard.html      ← Interfaz visual del dashboard
├── static/                 ← Imágenes, CSS extra (si se agregan)
└── data/                   ← Acá van tus archivos Excel
    └── .gitkeep
```

---

## ▶️ Cómo abrirlo

### Opción 1 — Doble-click (más fácil)
- **Windows:** doble-click en `INICIAR.bat`
- **Mac/Linux:** ejecutar `./iniciar.sh` en la terminal

### Opción 2 — Desde Visual Studio Code
1. Abrí VS Code
2. Abrí la carpeta `fixy_dashboard` (File → Open Folder)
3. Abrí la terminal integrada (Ctrl+` o View → Terminal)
4. Ejecutá estos comandos:

```bash
# Primera vez: instalar dependencias
pip install -r requirements.txt

# Iniciar el servidor
python app.py
```

5. Abrí tu navegador en: **http://localhost:5000**

---

## 📂 Cómo cargar tus datos

Una vez que el dashboard esté corriendo:

1. Abrí el navegador en http://localhost:5000
2. Usá el botón de carga (drag & drop) para subir tus Excel:
   - `Tablero_Fixy_AAAA.xlsx`
   - `FACTURACION_AAAA_DD_MM_AA.xlsx`
   - `Fulfillment_AAAA.xlsx`

O podés copiar los archivos Excel directamente a la carpeta `/data` y recargar la página.

---

## 🔧 Extensión recomendada para VS Code

Para ver el proyecto más cómodo, instalá:
- **Python** (Microsoft) — resalta el código de `app.py`
- **Live Server** — si querés previsualizar HTML estático

---

## ❓ Problemas frecuentes

| Problema | Solución |
|---|---|
| "Python no encontrado" | Instalá Python desde python.org y marcá "Add to PATH" |
| Puerto 5000 en uso | Cambiá el puerto en app.py: `port=5001` |
| Error al instalar deps | Usá `pip3` en vez de `pip` |
