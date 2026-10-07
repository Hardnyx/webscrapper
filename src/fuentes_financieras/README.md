# fuentes_financieras

Biblioteca compartida de fuentes externas integrada en la arquitectura de `Automatizaciones`.

Ubicación canónica del paquete:

```text
Automatizaciones/
└── librerias/
    └── fuentes/
        └── src/
            └── fuentes_financieras/
```

La resolución del almacén sigue este orden:

1. `FINANCIAL_SOURCES_DATA_ROOT`, cuando existe configuración explícita.
2. Ruta de datos declarada en `estructura.json`, cuando existe.
3. Almacén central poblado `Automatizaciones/datos/fuentes`.
4. Almacén interno ya poblado `librerias/fuentes/src/fuentes_financieras/data/sources`, para compatibilidad con instalaciones existentes.
5. `Automatizaciones/datos/fuentes` como ubicación inicial para instalaciones nuevas.

`AUTOMATIZACIONES_ROOT` fija la raíz del monorepo durante ejecuciones desde Colab u otros directorios de trabajo.

La API pública conserva `source()`, `fetch()`, `sync()`, `load()`, `describe()` y `list_datasets()`.

En `pe.smv.fondos_mutuos.valores_cuota`, la sincronización normal reutiliza particiones válidas existentes y completa períodos faltantes. El refresco forzado queda reservado para una ejecución explícita con `force=True`.
