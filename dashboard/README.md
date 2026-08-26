# Kampaoh Trueque — Dashboard

Dashboard web para gestionar el programa de intercambios artísticos de Kampaoh.  
Conecta en tiempo real con Google Sheets a través de Google Apps Script.

---

## Acceso rápido para el equipo

El dashboard está publicado en GitHub Pages.  
Solo necesitas la URL del Apps Script (la da la persona responsable del proyecto).

1. Abre el dashboard en el navegador
2. Pulsa **⚙ Configurar API** e introduce la URL del Apps Script
3. Pulsa **Guardar y conectar**

La URL se guarda en tu navegador. Solo hay que configurarla **una vez por dispositivo**.

---

## Estructura del repositorio

```
dashboard/
  index.html              ← El dashboard (un solo archivo HTML)
  README.md               ← Este archivo

KampaohTrueque_API.js     ← Apps Script: API del dashboard (correos, datos)
KampaohTrueque_Script.js  ← Apps Script: lógica de PLs y borradores
Manual_Trueque_Kampaoh.rtf ← Manual de usuario (Word editable)
```

---

## Configuración inicial (una vez por proyecto)

### Paso 1 — Subir el código a Google Apps Script

1. Abre el Google Sheet de BDD_Trueque
2. **Extensiones → Apps Script**
3. Crea dos archivos nuevos (`+` → Script):
   - `API` → pega el contenido de `KampaohTrueque_API.js`
   - `Script` → pega el contenido de `KampaohTrueque_Script.js`
4. Guarda (Ctrl+S)

### Paso 2 — Desplegar como Web App

1. En Apps Script: **Implementar → Nueva implementación**
2. Tipo: **Aplicación web**
3. Configuración:
   - **Ejecutar como:** Usuario que accede a la aplicación web
   - **Quién tiene acceso:** Cualquier usuario de Kampaoh (tu organización de Google Workspace)
4. Haz clic en **Implementar** y copia la URL generada

> La URL tiene este formato:  
> `https://script.google.com/macros/s/AKfycb.../exec`

### Paso 3 — Dar acceso al equipo

Comparte el Google Sheet con todos los miembros del equipo (acceso de Editor).  
Cada persona que acceda al dashboard por primera vez verá un popup de autorización de Google — es normal, solo hay que aceptarlo una vez.

---

## Configuración de correo desde marketing@kampaoh.com

Todos los borradores de correo se crean **desde** `marketing@kampaoh.com`.  
Para que esto funcione, **cada miembro del equipo** debe hacer esto una vez:

### Añadir marketing@kampaoh.com como alias en Gmail

1. Abre Gmail → **Configuración** (⚙) → **Ver toda la configuración**
2. Pestaña **Cuentas e importación** → sección **Enviar correo como**
3. Haz clic en **Añadir otra dirección de correo electrónico**
4. Nombre: `Kampaoh Trueque` / Email: `marketing@kampaoh.com`
5. Desactiva la casilla "Tratar como un alias"
6. Haz clic en **Siguiente paso** → llegará un código al grupo marketing@kampaoh.com
7. Copia el código de verificación y pégalo

A partir de ahí, cuando el dashboard genere un borrador, aparecerá en tus Borradores de Gmail listo para enviarse **como** marketing@kampaoh.com.

---

## Publicar en GitHub Pages

```bash
git init
git add .
git commit -m "feat: Kampaoh Trueque dashboard"
git remote add origin https://github.com/TU_USUARIO/kampaoh-trueque.git
git push -u origin main
```

Luego en GitHub:
- **Settings → Pages**
- Source: `Deploy from a branch` → Branch: `main` / folder: `/dashboard`

URL resultante: `https://tu-usuario.github.io/kampaoh-trueque/`

---

## Actualizar el código en Apps Script

Cada vez que se modifique `KampaohTrueque_API.js` o `KampaohTrueque_Script.js`:

1. Copia el contenido del archivo actualizado
2. Pégalo en el archivo correspondiente de Apps Script
3. **Implementar → Administrar implementaciones** → lápiz ✏️ → **Nueva versión** → **Implementar**

La URL no cambia. El dashboard coge la nueva versión automáticamente.

---

## Correos del programa (PLs)

| Código | Cuándo se envía | A quién |
|--------|----------------|---------|
| PL-02 | Primer contacto con el artista | Al artista |
| PL-05 | El artista acepta — condiciones definitivas | Al artista |
| PL-06 | Confirmación final (incluye localizador) | Al artista |
| PL-07 | Recordatorio antes de la llegada | Al artista |
| PL-08 | Aviso de llegada al camping | Al camping |
| PL-09 | Encuesta de valoración (post-trueque) | Al artista |

> **Regla crítica:** Todos los PLs se crean como **borradores** en Gmail.  
> Nunca se envían automáticamente. Siempre revisa el borrador antes de enviarlo.
