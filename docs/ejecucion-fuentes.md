# Ejecución conjunta de fuentes

`sync_fuentes.py` ejecuta un perfil TOML de proveedores independientes. Todas las consultas se planifican antes de la primera descarga, con un límite total de solicitudes (1000 por defecto). Los trabajos se ejecutan secuencialmente; un fallo de sincronización queda registrado y los demás trabajos continúan. No se resuelven dependencias entre fuentes ni se infieren cruces de entidades.

Un perfil reutilizable puede contener:

```toml
version = 1

[[trabajos]]
nombre = "Rentabilidad mensual"
dataset = "pe.sbs.rentabilidad"
[trabajos.consulta]
desde = "2026-07"
hasta = "2026-08"
tipos = ["B", "F", "C", "R"]
[trabajos.opciones]
keep_raw = true

[[trabajos]]
nombre = "Clasificaciones semestrales"
dataset = "pe.sbs.clasificaciones_riesgo"
[trabajos.consulta]
periodos = ["202601"]
[trabajos.opciones]
refresh_hours = 24
```

Las consultas usan exactamente los parámetros de cada proveedor; confirme los nombres y formatos de período en su documentación. Las opciones admitidas son `force`, `keep_raw` y `refresh_hours` positivo. No se permite habilitar cambios de esquema desde el perfil. Un nombre de dataset desconocido, perfil inválido, plan vacío o límite excedido cancela la ejecución antes de sincronizar fuentes.

```bash
python scripts/sync_fuentes.py --perfil mi_perfil.toml --data-root /ruta/datos --output-dir /ruta/reportes --solo-plan
python scripts/sync_fuentes.py --perfil mi_perfil.toml --data-root /ruta/datos --output-dir /ruta/reportes
python scripts/sync_fuentes.py --perfil mi_perfil.toml --data-root /ruta/datos --output-dir /ruta/reportes --solo-cache
```

`--solo-plan` no captura ni verifica caché. Para clasificaciones y su inventario se requieren `periodos` explícitos en los modos offline; no se descubre la última publicación y se aplica una comprobación conservadora de antigüedad a todos los cortes solicitados. La planificación normal de esos proveedores sí puede consultar el catálogo web antes de sincronizar. `--solo-cache` comprueba únicamente las capturas solicitadas, sin sincronización ni solicitudes de red. Ambos modos son excluyentes. El lanzador verifica e instala dependencias faltantes; esa preparación puede requerir internet incluso en modo de caché. El comando instalado `fuentes-sync` utiliza las dependencias ya instaladas.

Cada ejecución crea una carpeta única con `ejecucion.xlsx` y `ejecucion.json`. El Excel tiene hojas `trabajos` y `capturas`, encabezados españoles, tablas con filtros, estilo claro 9 y panel congelado. El JSON conserva estados para consumidores y programadores externos. No se sobrescriben reportes anteriores.

La verificación exige manifest validado, contrato y versión de parser vigentes, hash de la captura canónica coincidente y, para proveedores PDF que ofrecen verificación, archivo PDF íntegro. Para fuentes mutables también se comprueba si corresponde actualizar según `refresh_hours`. Un parser antiguo queda pendiente aunque la sincronización genérica haya omitido su caché; `force=true` permite reprocesarlo según el proveedor. Un período no disponible sigue pendiente incluso cuando se omite por caché. Una actualización fallida no se declara exitosa por conservar un archivo previo.

Los códigos de salida son 0 para selección completa (o plan generado), 2 para capturas pendientes y 1 para fallos. «Completo para la selección» describe la integridad y actualización de las capturas solicitadas; no garantiza que todos los campos financieros estén extraídos, que todas las entidades existan ni que la historia de la fuente esté completa. Cuando los metadatos del proveedor lo permiten, se muestra por separado el número de campos pendientes de revisión.

Este comando puede invocarse desde el Programador de tareas de Windows, cron o un flujo externo. Todavía no instala un calendario, descubre documentos nuevos, emite notificaciones ni mantiene un registro de novedades entre ejecuciones. Ejecute un único proceso escritor por almacén: la serialización de trabajos dentro del comando no bloquea escritores externos.

Validación con proveedores reales del repositorio: perfil de inventario semestral e informe de Huancayo en planificación y verificación offline, y perfil de PDF de Huancayo/Interbank en sincronización repetida. Se deshabilitó el transporte: los tres modos no hicieron solicitudes de red en estas consultas. La verificación marcó el inventario como pendiente de actualización por superar 24 horas; conservó el PDF como captura validada y mostró sus 12 campos pendientes de revisión. La sincronización repetida omitió ambos PDF por caché. Se comprobaron los reportes Excel y JSON fuera del repositorio.
