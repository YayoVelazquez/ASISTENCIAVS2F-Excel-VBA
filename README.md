# ASISTENCIAVS2F — Excel VBA (Ausentismo y Asistencia)

Este repositorio contiene el **código VBA exportado** desde el archivo Excel `.xlsm` para automatizar el cálculo y visualización de **asistencias/ausentismo** por mes.

## Qué hace el proyecto
- Lee códigos diarios de una hoja fuente (por defecto: `VSM2`).
- Construye/actualiza la tabla mensual `ConteoTbl` en hojas por mes.
- Calcula agregados y métricas (incluye indicadores de ausentismo operativo/administrativo según reglas implementadas en VBA/Excel).
- Crea/actualiza gráficos (ausencias, control, operativo+target, composición).

## Hojas principales (según el archivo original)
- `_CTRL`
- `INSTRUCCIONES`
- `VSM2` (fuente)
- `AGOSTO_2025`, `SEPTIEMBRE_2025`, `OCTUBRE_2025`, `NOVIEMBRE_2025`, `DICIEMBRE_2025`

## Estructura del repositorio
- `src/` → módulos VBA exportados (`.bas`, `.cls`)
- `docs/` → documentación/capturas (agrega aquí manuales o screenshots)

## Importante (seguridad)
- **No subas contraseñas** ni rutas internas.  
  En el archivo original se detectó al menos una constante de contraseña (p.ej. `PROTECT_PASSWORD`). Para publicación:
  - Sustituye contraseñas por placeholders.
  - Usa `Config_TEMPLATE.bas` como plantilla y excluye `Config.bas` real con `.gitignore`.

## Cómo reconstruir el .xlsm desde el repo (manual)
1. Abre Excel → `ALT+F11` (VBA Editor).
2. Importa archivos desde `src/`:
   - `File > Import File...` para `.bas` y `.cls`
3. Ajusta la configuración local (contraseñas/hoja fuente) en un módulo `Config.bas` (no versionado).

## Recomendación para GitHub
- **Repo público (portafolio):** sube `src/` + documentación + un `.xlsm` *sanitizado* (con datos dummy).
- **Repo privado (uso real):** puedes subir el `.xlsm` completo, además de `src/`.

---
Autor (según encabezados del código): Yael Velázquez Artolozaga  
Última exportación: 2025-12-15
