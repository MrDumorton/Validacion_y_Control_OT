# App Streamlit - Validación de Órdenes de Trabajo

Aplicación para revisar los campos esenciales definidos en las Órdenes de Trabajo (OT) y generar un reporte ejecutivo de faltantes y valores inválidos.

## Campos validados

### Anverso

- Horómetro.
- Motivo de detención del equipo.
- En cada fila utilizada del bloque **Información del trabajo**:
  - Descripción del síntoma.
  - Código trabajo.
  - Código síntoma.
  - Código causa.
- Descripción de actividades.

### Reverso

- Firma del jefe de turno: nombre y RUT.
- Firma del técnico responsable: nombre y RUT.

No se genera ninguna observación por los demás campos del anverso o reverso.

## Reglas especiales

- Los códigos causa **6.6** y **7.1** se consideran inválidos porque corresponden a la categoría **Otros**.
- La descripción de actividades se considera completa cuando existe información en al menos una de las líneas del bloque.
- En el resumen por OT se incluye la columna **Campos con observación**, inmediatamente después de **Estado**, con el detalle de los campos faltantes o inválidos.

## Campos críticos

- Código trabajo.
- Código síntoma.
- Código causa.
- Firma jefe turno (nombre + RUT).
- Firma técnico responsable (nombre + RUT).

El horómetro, el motivo de detención, la descripción del síntoma y la descripción de actividades se clasifican como campos estándar.

La aplicación muestra KPIs, un gráfico de cumplimiento por campo, estado de las OT, resumen por campo y detalle de hallazgos. También permite descargar reportes en Excel y PDF.

## Ejecución local recomendada

Usar Python 3.12:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Despliegue en Streamlit Community Cloud

1. Crear o actualizar el repositorio en GitHub.
2. Subir `app.py`, `requirements.txt` y este `README.md`.
3. En Streamlit Community Cloud seleccionar el repositorio.
4. Definir `app.py` como archivo principal.
5. Seleccionar Python 3.12 en las opciones avanzadas y desplegar.

## Coordenadas del formato

Las coordenadas se encuentran en `validate_work_order()`:

- Horómetro: `G13`.
- Motivo de detención: `AB25`.
- Filas de información del trabajo: desde la fila 42, cada cuatro filas.
- Código trabajo: columna `W`.
- Descripción del síntoma: columna `Z`.
- Código síntoma: columna `AO`.
- Código causa: columna `BH`.
- Descripción de actividades: bloque `B98:BZ124`.
- Firma jefe de turno: `C238` y `C244`.
- Firma técnico responsable: `BD239` y `BD243`.

## Módulo Control OT Semanal con Supabase

La aplicación incorpora el selector superior **Validación OT / Control OT Semanal**.

El nuevo módulo permite:

- Importar el reporte de detenciones desde la hoja `Limpio`.
- Excluir automáticamente categorías configurables, por ejemplo lavado, neumáticos, automatización y cámaras.
- Guardar en Supabase las OT digitales procesadas previamente por el módulo de validación.
- Asociar automáticamente una OT con una detención utilizando equipo y similitud de descripción.
- Confirmar manualmente asociaciones cuando sea necesario.
- Asignar una OT recibida posteriormente al día original de la detención.
- Mostrar cumplimiento porcentual diario.
- Mostrar debajo el detalle de las OT faltantes con fecha, hora, turno, equipo, razón y comentario.
- Evitar duplicados mediante identificadores únicos de detenciones y números OT únicos.

### Configuración de Supabase

1. Crear un proyecto gratuito en Supabase.
2. Abrir **SQL Editor** y ejecutar el contenido de `supabase_schema.sql`.
3. En Streamlit Community Cloud abrir **Settings > Secrets**.
4. Agregar:

```toml
[supabase]
url = "https://TU-PROYECTO.supabase.co"
secret_key = "TU_SECRET_KEY"
```

La `secret_key` (empieza con `sb_secret_`) debe mantenerse únicamente en Streamlit Secrets y nunca subirse al repositorio público. No uses la publishable key para este módulo, porque las tablas tienen RLS habilitado y la aplicación necesita operaciones de servidor.

### Flujo diario recomendado

1. Entrar en **Control OT Semanal** e importar el reporte de detenciones.
2. Entrar en **Validación OT**, subir y validar las OT digitales.
3. Volver a **Control OT Semanal** y presionar **Guardar OT validadas**.
4. Revisar el resumen diario y confirmar manualmente las asociaciones que no sean evidentes.

Una detención puede tener varias OT relacionadas. Para el indicador diario, una detención se considera respaldada cuando tiene al menos una OT digital confirmada.

## Corrección de asociación por equipo y horario

La asociación automática ahora normaliza códigos como `TO-28_D10T2` a `TO28` y
`MO-10_2` a `MO10`. Además, lee desde la OT las fechas y horas de inicio/término
(`X9`, `AB9`, `X11`, `AB11`) y las compara contra el intervalo real de la detención.
Los valores mostrados en AM/PM y los horarios de 24 horas se convierten al mismo
objeto de fecha/hora antes de comparar.

Para una base Supabase ya creada, ejecutar una vez en SQL Editor:

```text
supabase_migration_matching.sql
```

Después se deben volver a importar las detenciones y volver a guardar las OT para
completar los nuevos campos de equipo normalizado e intervalo operacional.
