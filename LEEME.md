# 🐎 El Patrón — Módulo de Fichaje (Google Sheets)

Sistema web independiente de control de horarios y asistencia de empleados para **El Patrón**, diseñado con la estética oficial del sistema principal y sincronizado con **Google Sheets + Google Apps Script**.

---

## 📁 Archivos incluidos en esta carpeta

1. **`index.html`**: Aplicación web completa en un solo archivo (HTML + CSS + JS nativo, sin frameworks).
2. **`Code.gs`**: Backend para Google Apps Script (vinculado a una planilla de Google Sheets).
3. **`LEEME.md`**: Esta guía paso a paso de instalación y puesta en marcha.
4. *(Opcional)* **`logo.png`**: Si ponés el archivo `logo.png` en esta misma carpeta junto a `index.html`, se mostrará automáticamente en el sidebar, en el login y en el banner superior. Si no está el archivo, el sistema muestra automáticamente el ícono `🐎`.

---

## 🚀 Paso 1: Configurar la Planilla de Google Sheets y el Backend (`Code.gs`)

1. Entrá a [Google Sheets](https://sheets.google.com) con tu cuenta de Google y creá una **Hoja de cálculo en blanco**.
2. Ponele de nombre, por ejemplo: **`El Patrón - Fichaje de Personal`**.
3. En el menú superior de la planilla, andá a **Extensiones > Apps Script**.
4. Se va a abrir el editor de código. Borrá cualquier texto que venga por defecto en `Código.gs` (como `function myFunction() {}`) y **pegá todo el contenido del archivo `Code.gs`** de esta carpeta.
5. *(Opcional)* Arriba del todo en `Code.gs` vas a ver la constante `SALT`. Podés dejarla como está o cambiarla por la frase secreta que prefieras **antes** de crear los usuarios.
6. Tocá el ícono del disquete **Guardar proyecto** (o presioná `Ctrl + S`).

---

## ⚙️ Paso 2: Inicializar las hojas y el usuario Dueño (`setup`)

1. En la barra superior del editor de Apps Script, al lado del botón **Ejecutar**, desplegá el selector de funciones y elegí **`setup`**.
2. Hacé clic en **Ejecutar**.
3. Google te va a pedir **Revisar permisos**:
   - Elegí tu cuenta de Google.
   - Hacé clic en **Configuración avanzada** (abajo a la izquierda).
   - Hacé clic en **Ir a Proyecto sin título (no seguro)** y luego en **Permitir**.
4. Cuando termine la ejecución, volvé a tu planilla de Google Sheets y vas a ver que se crearon automáticamente dos pestañas con el encabezado cobre/negro de El Patrón:
   - **`Usuarios`** (`Usuario`, `Nombre`, `Rol`, `PIN hash`, `Activo`, `Creado`) con el usuario inicial ya cargado.
   - **`Fichajes`** (`ID`, `Usuario`, `Nombre`, `Fecha`, `Ingreso`, `Egreso`, `Horas decimal`, `Duración hh:mm`, `Estado`, `IngresoTS`, `EgresoTS`).

### 🔑 Credenciales iniciales creadas por `setup()`:
- **Usuario:** `dueno`
- **PIN:** `1234`
- **Rol:** `dueño` *(Te recomendamos cambiar el PIN desde la pestaña **"Mi cuenta"** apenas ingreses por primera vez).*

---

## 🌐 Paso 3: Publicar como Aplicación Web y obtener la URL `/exec`

1. Dentro del editor de Apps Script, hacé clic en el botón azul **Implementar** (arriba a la derecha) > **Nueva implementación**.
2. Al lado de "Seleccionar tipo", tocá el engranaje ⚙️ y elegí **Aplicación web**.
3. Completá los campos así:
   - **Descripción:** `API Fichaje El Patrón v1`
   - **Ejecutar como:** **`Yo`** *(tu cuenta de Google)*
   - **Quién tiene acceso:** **`Cualquier persona`** *(fundamental para que la web pueda conectarse sin bloqueos de CORS)*
4. Hacé clic en **Implementar**.
5. Copiá la **URL de la aplicación web** (empieza con `https://script.google.com/macros/s/...` y termina en **`/exec`**).

> **Nota importante:** Si más adelante modificás algo en `Code.gs`, acordate de ir a **Implementar > Administrar implementaciones > Editar (lápiz) > Versión: Nueva versión > Implementar** para que los cambios impacten en la URL `/exec`.

---

## 💻 Paso 4: Conectar y abrir el Frontend (`index.html`)

Tenés **dos maneras** de configurar la URL `/exec` en el frontend:

### Opción A: Desde la pantalla de Login (sin tocar código)
1. Abrí `index.html`.
2. Abajo del botón de ingreso, hacé clic en **`⚙️ Configurar conexión`**.
3. Pegá la URL de Apps Script (la que termina en `/exec`) y tocá **Guardar conexión**. Queda guardada en el `localStorage` de ese navegador/terminal.

### Opción B: Dejarla fija en el código (para subir a un hosting)
1. Abrí `index.html` con cualquier editor de texto (Bloc de notas, VS Code, etc.).
2. Buscá al inicio de la etiqueta `<script>` la línea:
   ```javascript
   const API_URL = "";
   ```
3. Pegá tu URL entre las comillas:
   ```javascript
   const API_URL = "https://script.google.com/macros/s/TU_ID_AQUI/exec";
   ```
4. Guardá el archivo. Al tener `API_URL` definida en el código, ya no se muestra el enlace de configuración en el login y todos los dispositivos se conectan directo a tu planilla.

---

## 🖥️ Cómo abrir el proyecto en local

El archivo `index.html` no usa ningún framework ni requiere compilación:

1. **Doble clic directo:** Podés abrir `index.html` haciendo doble clic desde el explorador de archivos.
2. **Con servidor local:** Si tenés Node.js instalado, abrí una terminal en esta carpeta y ejecutá:
   ```bash
   npx serve .
   ```
   Luego abrí la dirección local que te indique en consola (por ejemplo `http://localhost:3000`).

> 🧪 **Modo Local de Prueba Integrado:** Si abrís `index.html` en local sin haber pegado todavía una URL `/exec`, el sistema funciona automáticamente con un motor local idéntico a `Code.gs` (con el usuario inicial `dueno` y PIN `1234`), para que puedas probar todas las pantallas, cronómetros, altas de empleados y descargas de CSV al instante.

---

## 🛡️ Resumen de Funciones y Seguridad

- **PIN encriptado:** Nunca se guarda el PIN en texto plano en la Google Sheet; se almacena un hash **SHA-256 con SALT**.
- **Protección contra fuerza bruta:** Tras **5 intentos fallidos** de login, la cuenta se bloquea por **5 minutos** tanto en el servidor (`CacheService`) como en la pantalla.
- **Anti doble fichaje:** Si un empleado ya marcó **INGRESO**, el botón **INGRESO** queda deshabilitado hasta que marque **EGRESO**, y el servidor valida con `LockService` para impedir fichajes duplicados por doble clic o doble pestaña.
- **Sin GPS:** No solicita ni rastrea ubicación del dispositivo en ningún momento.
- **Zona horaria fija:** `America/Argentina/Cordoba` en el reloj en vivo, en los filtros y en el servidor.
- **Exportación para Excel argentino:** El botón **Descargar CSV** de la pestaña **Reporte** incluye BOM UTF-8 (para tildes y eñes), separador `;` y decimales con coma `,`.
