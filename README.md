# Longo Finance · Gestión de carteras

Página web de carteras (Portfolio, Alertas e Informes) servida por Google Apps Script sobre la planilla
"Carteras en vivo". Cada cambio que se sube a `main` dentro de `apps-script/` se publica solo en la misma
implementación, así que el link no cambia.

## Estructura
- `apps-script/Codigo.gs` — precios (Rava), alertas de cheques, informes premarket, postmarket y semanal, y datos para la página.
- `apps-script/Index.html` — la página.
- `apps-script/appsscript.json` — configuración del proyecto (zona horaria y acceso a la app web).
- `.github/workflows/deploy.yml` — publicación automática.

## Configuración (una sola vez)
1. En https://script.google.com/home/usersettings activar **API de Google Apps Script**.
2. En la Terminal de la Mac (requiere Node.js):
   ```bash
   npx @google/clasp@2.4.2 login
   ```
   Se abre el navegador: elegir la cuenta dueña de la planilla y aceptar. Después copiar el archivo de credenciales:
   ```bash
   cat ~/.clasprc.json | pbcopy
   ```
3. En GitHub → repositorio → **Settings → Secrets and variables → Actions**:
   - pestaña **Secrets** → *New repository secret*: nombre `CLASPRC_JSON`, valor: pegar lo copiado.
   - pestaña **Variables** → *New repository variable*:
     - `SCRIPT_ID`: Apps Script → ⚙ Configuración del proyecto → ID de la secuencia de comandos.
     - `DEPLOYMENT_ID`: Apps Script → Implementar → Gestionar implementaciones → ID de la implementación de la app web.
4. GitHub → **Actions → Publicar en Apps Script → Run workflow** para probar.

## Notas
- La planilla, los disparadores y los permisos de Google no se tocan con la publicación: solo cambia el código.
- Si se agregan funciones nuevas con disparadores, volver a correr **Carteras → Instalar automatizaciones** desde la planilla.
