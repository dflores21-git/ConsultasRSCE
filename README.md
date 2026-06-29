# RSCE — App de Consultas SQL

## Requisitos previos
- Node.js instalado (versión 18 o superior recomendada) — https://nodejs.org
- Acceso de red al servidor SQL Server `RSCEDATOS01`

## Instalación

1. Descarga/copia esta carpeta a tu equipo.
2. Abre una terminal dentro de la carpeta y ejecuta:
   ```bash
   npm install
   ```
   Esto instalará: express, mssql, exceljs, pdfkit y dotenv.

3. Crea tu archivo de configuración:
   ```bash
   cp .env.example .env
   ```
   Edita `.env` y comprueba que los valores son correctos para tu entorno.
   **No subas `.env` a ningún repositorio ni lo compartas** — contiene la contraseña de la base de datos.

## Ejecutar la aplicación

```bash
npm start
```

o directamente:

```bash
node app.js
```

Verás en la terminal:
```
✅  Aplicación RSCE iniciada
   Abre en tu navegador: http://localhost:3000
```

## Importante: no hay interfaz visual incluida

El código original solo define el **backend** (las consultas SQL y los endpoints `/api/...`). No incluye una página web (HTML/botones) para usarlo desde el navegador — eso estaría en una carpeta `public/` que no venía en el documento original.

Por ahora puedes probar los endpoints directamente, por ejemplo con `curl` o Postman:

```bash
# Listar razas disponibles
curl http://localhost:3000/api/razas

# Ejecutar la consulta de socios (sin filtros)
curl -X POST http://localhost:3000/api/query/socios -H "Content-Type: application/json" -d "{}"

# Exportar a Excel
curl -X POST http://localhost:3000/api/export/excel/socios -H "Content-Type: application/json" -d "{}" -o socios.xlsx
```

Si quieres, puedo crearte también una página web sencilla (`public/index.html`) con un desplegable para elegir la consulta, filtros y botones de exportar a Excel/PDF — dímelo y la preparo.

## Seguridad

- Las credenciales de la base de datos ya NO están escritas en `app.js`; se leen desde `.env`.
- Te recomiendo cambiar la contraseña de la base de datos (`MastinEspanol987&`), ya que ha estado en texto plano en un documento Word.
