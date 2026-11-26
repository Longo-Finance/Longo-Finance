# Longo Finance — instrucciones para Claude

- El dueño del repositorio pidió que **todo cambio que solicite quede publicado en la app automáticamente**:
  una vez hecho y verificado, integrarlo y subirlo a `main` (además de la rama de trabajo). El push a `main`
  dispara `.github/workflows/deploy.yml`, que publica en Apps Script con el mismo link.
- Después de subir a `main`, revisar que la ejecución de "Publicar en Apps Script" termine bien y avisar si falla.
- Si se cambian disparadores (horarios) en `apps-script/Codigo.gs`, subir `VERSION_DISPARADORES`: el disparador de
  precios los reinstala solo en menos de 15 minutos, sin pasar por el menú de la planilla.
- Textos de la app y de los informes en español rioplatense formal.
