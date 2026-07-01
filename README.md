# Libro de Conciliación de Guías — BravoEnergy / Cementos Melón

Herramienta web para conciliar números de guía de despacho entre 3 archivos Excel
(Laboratorio, JD Edwards y Proveedor) al preparar el estado de pago mensual de BravoEnergy.

Es una aplicación 100% estática (HTML + CSS + JS, sin backend). Todo el procesamiento
de los archivos Excel ocurre **en el navegador**, con [SheetJS](https://sheetjs.com/);
ningún archivo se sube a ningún servidor.

## Cómo ejecutarla localmente

Los navegadores modernos bloquean la lectura de archivos locales (`file://`) por
seguridad en algunos casos, así que la forma más confiable de abrir la app es con un
servidor local mínimo:

```bash
# Python 3 (ya viene instalado en la mayoría de los sistemas)
python -m http.server 8080
```

Luego abre `http://localhost:8080` en tu navegador.

Alternativamente, en muchos navegadores (Chrome, Edge) también puedes simplemente
hacer doble clic en `index.html` y funcionará sin problemas. Si ves errores al cargar
los archivos Excel, usa el método del servidor local de arriba.

## Uso

1. Sube los 3 archivos Excel en sus respectivos casilleros (Laboratorio, JD Edwards, Proveedor).
2. En "Mapeo de columnas", confirma o corrige cuál columna de cada archivo contiene el número de guía.
3. Haz clic en "Conciliar guías".
4. Revisa el resumen, filtra/busca en la tabla, y descarga el Excel de resultados (hojas "Resumen" y "Discrepancias").

## Estructura del proyecto

- `index.html` — estructura de la página.
- `style.css` — estilos (paleta, tipografía, layout tipo "libro contable").
- `app.js` — lógica de carga, normalización de guías, conciliación 3-way, filtros y exportación.

## Privacidad

Ningún dato se envía a un servidor externo. Todo el procesamiento (lectura de Excel,
comparación y generación del archivo de resultados) ocurre localmente en el navegador
de quien use la herramienta. Si publicas este repositorio como sitio público (GitHub
Pages), cualquiera con el enlace podrá usar la herramienta, pero los archivos Excel
que suban seguirán procesándose únicamente en su propio navegador — no en ningún
servidor tuyo ni de terceros.
