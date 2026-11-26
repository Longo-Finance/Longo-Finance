# Longo Finance · Gestión de carteras

Página web de carteras (Portfolio, Alertas e Informes) servida por Google Apps Script sobre la planilla
"Carteras en vivo". Cada cambio que se sube a `main` dentro de `apps-script/` se publica solo en la misma
implementación, así que el link no cambia.

## Accesos (portal de clientes)
La app web se ejecuta con la cuenta del dueño y no muestra ningún dato sin **link personal + clave**.
La pestaña **Accesos** de la planilla (se crea sola) tiene una fila por cliente con sus cuentas y una de
**Administración** (todas las cuentas, Alertas, Informes y "Ver como cliente").
- Link personal: columna *Link personal*. Clave: columna *Clave* (6 dígitos).
- Cambiar una clave: borrarla; en el próximo ciclo (≤15 min) o con *Carteras → Accesos de clientes* se genera otra.
- Revocar un link: borrar el *Token* (se genera uno nuevo y el anterior deja de funcionar).
- Desactivar un acceso: poner `NO` en *Activo*.
- Agregar otro administrador o cliente: nueva fila con Nombre, Rol (`admin`/`cliente`) y Cuentas (separadas por coma).
- 5 claves erróneas seguidas bloquean ese link 15 minutos. Las sesiones duran 6 horas desde el último uso.

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

## Informes premarket y postmarket (estilo diario)
Cada informe arma una portada (titular, bajada, claves del día y mapa de mercados), las noticias del día con resumen e imagen
(Ámbito, Infobae, El Cronista, La Nación, CNBC, MarketWatch, Yahoo Finance y Google Noticias para las empresas en cartera),
gráficos de ADRs y de tendencia de riesgo país y dólar MEP, y los datos de mercado y de la cartera.

Horarios (lunes a viernes, hora Argentina): el **premarket** se regenera cada 30 minutos entre las 8:00 y las 11:30 y el
**postmarket** cada 30 minutos entre las 18:00 y las 20:00. Fuera de esas franjas queda la última edición hasta el día siguiente
(desde el menú Carteras de la planilla se puede regenerar a mano).

Calendario económico: Forex Factory (EE.UU., zona euro, China, Reino Unido y Japón; impacto alto y medio) más los eventos
locales que se carguen en la pestaña **Agenda** de la planilla (Fecha, Hora, Evento, Importancia), por ejemplo licitaciones
del Tesoro, datos del INDEC o reuniones del BCRA.

Dos agregados opcionales, que se activan en Apps Script → ⚙ Configuración del proyecto → **Propiedades de la secuencia de comandos**:
- `ANTHROPIC_API_KEY`: la portada la redacta Claude a partir de los datos y titulares del informe (sin datos de clientes).
  Sin la clave, la portada se arma con reglas fijas.
- `TELEGRAM_CANALES`: canales públicos de Telegram que se leen sin clave para "Lo que se dice en redes"
  (por defecto `finanzasargy`; lista separada por comas).
- `X_BEARER_TOKEN`: suma la sección "Lo que se dice en X" con los posteos más relevantes de las cuentas seguidas
  (requiere acceso a la API de X, que es pago). Las cuentas se cambian con la propiedad `X_CUENTAS`
  (lista separada por comas, por ejemplo `BCRAoficial,LuisCaputoAR,WSJmarkets`).

## Notas
- La planilla, los disparadores y los permisos de Google no se tocan con la publicación: solo cambia el código.
- Los disparadores se actualizan solos: si el código cambia horarios (constante `VERSION_DISPARADORES` en `Codigo.gs`),
  el disparador de precios los reinstala en menos de 15 minutos. **Carteras → Instalar automatizaciones** sigue disponible
  para la instalación inicial o para forzarlo.
