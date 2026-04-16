# Generador de Boletines Web

Aplicacion web en Flask para:
- Cargar archivo Excel (.xlsx)
- Seleccionar periodos (001, 002, 003, FINAL)
- Generar y descargar reporte final en Excel

## Ejecutar en local (Windows PowerShell)

```powershell
& .\.venv\Scripts\Activate.ps1
pip install -r requirements-web.txt
flask --app web_app run --host 0.0.0.0 --port 5000
```

Abrir en navegador:
http://127.0.0.1:5000

## Despliegue con enlace publico (Render)

1. Sube este proyecto a GitHub.
2. En Render, crea un nuevo `Web Service` desde ese repositorio.
3. Render detectara `render.yaml` y aplicara:
   - Build: `pip install -r requirements-web.txt`
   - Start: `gunicorn wsgi:app`
4. Al terminar, Render te entrega una URL publica.

Nota: en entorno web publico se prioriza la exportacion a Excel. Las
dependencias opcionales de PDF/Word pueden requerir configuracion adicional
en algunos proveedores.

## Despliegue en Railway

1. Conecta el repositorio en Railway.
2. Configura:
   - Build command: `pip install -r requirements-web.txt`
   - Start command: `gunicorn wsgi:app`
3. Railway publicara la app y te dara el enlace.

## Endpoint de salud

- `GET /health` retorna `{"status":"ok"}`.
