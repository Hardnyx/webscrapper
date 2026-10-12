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

Cada ejecución crea una carpeta única con `ejecucion.xlsx` y `ejecucion.json`. El Excel tiene hojas `trabajos`, `capturas` y `novedades`, encabezados españoles, tablas con filtros, estilo claro 9 y panel congelado. El JSON conserva estados para consumidores y programadores externos. No se sobrescriben reportes anteriores.

La verificación exige manifest validado, contrato y versión de parser vigentes, hash de la captura canónica coincidente y, para proveedores PDF que ofrecen verificación, archivo PDF íntegro. Para fuentes mutables también se comprueba si corresponde actualizar según `refresh_hours`. Un parser antiguo queda pendiente aunque la sincronización genérica haya omitido su caché; `force=true` permite reprocesarlo según el proveedor. Un período no disponible sigue pendiente incluso cuando se omite por caché. Una actualización fallida no se declara exitosa por conservar un archivo previo.

Los códigos de salida son 0 para selección completa (o plan generado), 2 para capturas pendientes y 1 para fallos. «Completo para la selección» describe la integridad y actualización de las capturas solicitadas; no garantiza que todos los campos financieros estén extraídos, que todas las entidades existan ni que la historia de la fuente esté completa. Cuando los metadatos del proveedor lo permiten, se muestra por separado el número de campos pendientes de revisión.

Este comando puede invocarse desde el Programador de tareas de Windows, cron o un flujo externo. Todavía no instala un calendario, no descubre documentos nuevos ni emite notificaciones. Ejecute un único proceso escritor por almacén: la serialización de trabajos dentro del comando no bloquea escritores externos.

Validación con proveedores reales del repositorio: perfil de inventario semestral e informe de Huancayo en planificación y verificación offline, y perfil de PDF de Huancayo/Interbank en sincronización repetida. Se deshabilitó el transporte: los tres modos no hicieron solicitudes de red en estas consultas. La verificación marcó el inventario como pendiente de actualización por superar 24 horas; conservó el PDF como captura validada y mostró sus 12 campos pendientes de revisión. La sincronización repetida omitió ambos PDF por caché. Se comprobaron los reportes Excel y JSON fuera del repositorio.


## Registro persistente de novedades

Desde la versión 0.28, cada ejecución compara las capturas verificadas con el último estado observado en `--data-root/state/novedades.sqlite3`. La identidad es dataset y período, independientemente del nombre del trabajo, perfil o carpeta de reportes. SQLite conserva el historial de transiciones con un identificador de evento, ejecución, fecha de detección UTC y huellas anterior/actual. Copie este archivo junto con la caché al trasladar el almacén; borrarlo reinicia la línea base.

La primera observación se etiqueta «Primera captura observada»: no demuestra que la publicación sea nueva. Las siguientes distinguen PDF cambiado, reprocesamiento o cambio de contrato, y contenido canónico cambiado. Se comparan el hash canónico, hash del PDF cuando está disponible, versión del parser y contrato. Si cambia el parser junto con el contenido, se informa como reprocesamiento, sin atribuirlo a la fuente. Los tiempos de consulta y el orden de las filas no generan novedades. Una reversión A → B → A constituye dos cambios reales; volver a observar A sin cambios no añade eventos.

Solo avanzan capturas validadas de trabajos sin fallo. Los períodos ausentes, no disponibles, desactualizados o dañados mantienen la observación anterior. Los períodos fuera del perfil no se interpretan como retirados. La comparación opera por captura completa, sin identificar todavía qué fila o indicador varió. Un PDF sin cambio de extracción puede producir una novedad documental; los proveedores sin hash del documento solo permiten observar su contenido canónico.

El Excel incluye la hoja `novedades`, incluso vacía. El JSON de ejecución tiene `version=2` y la lista `novedades`. `--solo-cache` puede establecer la primera observación y registrar cambios de capturas válidas sin consultar la fuente. `--solo-plan` no crea ni modifica el registro. Los códigos de salida conservan su significado y no dependen de cuántas novedades haya.

La transacción del registro se confirma después de escribir ambos reportes. Si falla la exportación, se revierte y la siguiente ejecución vuelve a mostrar las novedades. Un reporte parcial de una ejecución fallida no es confirmación de registro. La base SQLite serializa sus escritores, pero no bloquea la captura de proveedores externos: mantenga un único proceso escritor por almacén. Esto no constituye entrega de avisos ni instala un calendario; un consumidor podrá usar el identificador de evento para deduplicar sus propias notificaciones.

Validación del registro con los PDF reales de Huancayo e Interbank: dos primeras observaciones y cero novedades al repetir la verificación, con el transporte deshabilitado. Reportes Excel y JSON comprobados fuera del repositorio.
