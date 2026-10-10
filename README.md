# Turnos Olimpiadas Colegio Parroquial San Miguel

Publicar gratis, sin dominio y sin login.

Dos partes: una Google Sheet con un Apps Script (guarda los datos) y una página estática `index.html` (lo que ven las personas).

## 1. Crear la hoja y el backend (10 min)
1. En Google Drive crea una **Hoja de cálculo nueva** (vacía).
2. Menú **Extensiones > Apps Script**. Borra el contenido y pega todo `Code.gs`.
3. Cambia `CLAVE_ADMIN` (arriba del archivo) por una clave tuya.
4. En el editor elige la función **cargarDatos** y presiona **Ejecutar**. Acepta los permisos. Se crea la pestaña `Cupos` con los 88 cupos.
5. **Implementar > Nueva implementación > tipo "Aplicación web"**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier persona**
6. Copia la **URL de la aplicación web** (termina en `/exec`).

## 2. Poner la URL en la página
Abre `index.html` (en GitHub, botón del lápiz) y reemplaza `PEGA_AQUI_LA_URL_DE_APPS_SCRIPT` por esa URL.

## 3. Publicar gratis (GitHub Pages)
1. Crea un repositorio nuevo en GitHub (público) y sube `index.html`.
2. **Settings > Pages > Deploy from a branch > main / (root)**.
3. En un minuto queda en `https://TU-USUARIO.github.io/NOMBRE-DEL-REPO/`.

Alternativa sin cuenta de GitHub: arrastra la carpeta a https://app.netlify.com/drop.

## Uso
- Participantes: la URL normal. Cada pestaña (Apoderados / Delegados) lista los cupos. Lo guardado se ve oculto y no se puede modificar.
- Administración: la misma URL con `#admin` al final. Pide la clave. Permite editar, agregar, eliminar y descargar CSV.
- También puedes ver y editar todo directo en la Google Sheet.

## Notas
- Si cambias `Code.gs` después de implementar, usa **Implementar > Administrar implementaciones > Editar > Nueva versión**.
- Los nombres y cursos guardados no se envían a los participantes; solo ven que el cupo está ocupado. La vista `#admin` sí los ve.
- Cuentas gratuitas de Google tienen cuotas diarias de Apps Script. Para unos cientos de personas alcanza de sobra.
- Con la página abierta cualquiera puede inscribirse. No hay captcha ni verificación de identidad.
