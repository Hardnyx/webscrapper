"""Explicit consumer-level batch execution and requested capture coverage."""
from dataclasses import dataclass
from itertools import islice
import math
from pathlib import Path
import tomllib

from .registry import get_provider, get_spec
from .provider import DatasetProvider
from .storage import utc_now_iso

SYNC_OPTIONS = {'force', 'keep_raw', 'refresh_hours'}


@dataclass
class PlannedJob:
    name: str
    provider: object
    query: dict
    options: dict
    requests: list


def read_profile(path):
    """Validate the entire declarative profile before constructing providers."""
    path = Path(path)
    if path.stat().st_size > 1_000_000:
        raise ValueError('Perfil mayor de 1 MB.')
    data = tomllib.loads(path.read_text(encoding='utf-8'))
    if set(data) != {'version', 'trabajos'} or type(data['version']) is not int or data['version'] != 1:
        raise ValueError('Perfil: se requieren version=1 y trabajos, sin claves adicionales.')
    jobs = data['trabajos']
    if not isinstance(jobs, list) or not 1 <= len(jobs) <= 100:
        raise ValueError('El perfil debe contener entre 1 y 100 trabajos.')
    names = set()
    for job in jobs:
        if not isinstance(job, dict) or set(job) - {'nombre', 'dataset', 'consulta', 'opciones'}:
            raise ValueError('Trabajo con claves desconocidas.')
        name, dataset = job.get('nombre'), job.get('dataset')
        if not isinstance(name, str) or not name.strip() or name in names:
            raise ValueError('Cada trabajo requiere nombre único y no vacío.')
        names.add(name)
        if not isinstance(dataset, str):
            raise ValueError('Cada trabajo requiere dataset.')
        get_spec(dataset)
        query, options = job.get('consulta', {}), job.get('opciones', {})
        if not isinstance(query, dict) or not isinstance(options, dict):
            raise ValueError('consulta y opciones deben ser tablas TOML.')
        if set(options) - SYNC_OPTIONS or set(query) & (SYNC_OPTIONS | {'allow_schema_change'}):
            raise ValueError('Opciones no admitidas o mezcladas con consulta.')
        for key in ('force', 'keep_raw'):
            if key in options and type(options[key]) is not bool:
                raise ValueError(f'{key} debe ser booleano.')
        hours = options.get('refresh_hours', 24.0)
        if type(hours) not in (int, float) or not math.isfinite(hours) or hours <= 0:
            raise ValueError('refresh_hours debe ser positivo y finito.')
    return jobs


def plan_jobs(jobs, *, max_requests=1000, offline=False):
    """Plan every job before the first sync, with a total request bound."""
    if type(max_requests) is not int or max_requests < 1:
        raise ValueError('Límite de solicitudes inválido.')
    planned, total = [], 0
    for job in jobs:
        provider = get_provider(job['dataset'])
        query, options = job.get('consulta', {}), job.get('opciones', {})
        planner = provider.plan_offline if offline else provider.plan_sync
        requests = list(islice(planner(**query), max_requests - total + 1))
        total += len(requests)
        if not requests or total > max_requests:
            raise ValueError('Plan vacío o límite total de solicitudes excedido; no se sincroniza ningún trabajo.')
        keys = [r.period_key for r in requests]
        if len(set(keys)) != len(keys):
            raise ValueError('El proveedor generó solicitudes duplicadas.')
        planned.append(PlannedJob(job['nombre'], provider, query, options, requests))
    return planned


def capture_status(provider, request, refresh_hours):
    """Check requested canonical captures, not financial completeness or legal eligibility."""
    entry = provider.storage.manifest.get(request.period_key)
    if not entry:
        return 'Ausente', '', 0, 0
    checked = entry.get('checked_at', '')
    rows = entry.get('rows', 0)
    pending = sum(count for status, count in entry.get('metadata', {}).get('field_counts', {}).items()
                  if status != 'extracted')
    if entry.get('contract_version', '1') != provider.contract_version:
        status = 'Contrato incompatible'
    elif entry.get('parser_version') != provider.parser_version:
        status = 'Parser desactualizado; reprocesar con force'
    elif entry.get('status') == 'unavailable':
        status = 'No disponible según la fuente'
    elif entry.get('status') != 'validated':
        status = 'Sin validar'
    elif not provider.storage.period_matches(request.partition_key, request.period_key, entry.get('content_hash')):
        status = 'Captura canónica dañada o incompleta'
    elif hasattr(provider, 'pdf_matches') and not provider.pdf_matches(request.period_key, entry.get('metadata', {}).get('pdf_sha256')):
        status = 'PDF ausente o dañado'
    elif DatasetProvider._refresh_due(provider, request, entry, refresh_hours):
        status = 'Actualización pendiente'
    else:
        status = 'Captura validada'
    return status, checked, rows, pending


def execute_jobs(planned, *, cache_only=False, plan_only=False):
    if cache_only and plan_only:
        raise ValueError('solo-cache y solo-plan son excluyentes.')
    jobs, coverage = [], []
    for job in planned:
        started = utc_now_iso()
        result, error = None, ''
        if not cache_only and not plan_only:
            try:
                result = job.provider.sync(**job.options, **job.query)
                if result.requested != len(job.requests):
                    raise ValueError('El plan cambió durante la ejecución; revise la consulta.')
                if {d['period_key'] for d in result.details} != {r.period_key for r in job.requests}:
                    raise ValueError('El resultado no corresponde al plan solicitado.')
            except Exception as exc:
                error = f'{type(exc).__name__}: {exc}'
        details = {d['period_key']: d for d in result.details} if result else {}
        statuses = []
        for request in job.requests:
            status, checked, rows, pending = ('Planeada', '', 0, 0)
            if not plan_only:
                try:
                    status, checked, rows, pending = capture_status(job.provider, request, job.options.get('refresh_hours', 24.0))
                except Exception as exc:
                    status = 'Error al verificar caché'
                    error = error or f'{type(exc).__name__}: {exc}'
            detail = details.get(request.period_key, {})
            if detail.get('status') == 'failed':
                status = 'Falló actualización; revisar captura previa'
            statuses.append(status)
            coverage.append({'Trabajo': job.name, 'Dataset': job.provider.spec.dataset_id,
                'Período solicitado': request.period_key, 'Partición': request.partition_key,
                'Estado de captura': status, 'Última comprobación': checked,
                'Filas guardadas': rows, 'Campos pendientes de revisión': pending,
                'Resultado del proveedor': detail.get('status', ''), 'Error': detail.get('error', '')})
        complete = all(status == 'Captura validada' for status in statuses)
        failed = bool(error or (result and result.failed))
        jobs.append({'Trabajo': job.name, 'Dataset': job.provider.spec.dataset_id,
            'Estado': 'Planeado' if plan_only else 'Falló' if failed else 'Completo para la selección' if complete else 'Pendiente',
            'Solicitudes': len(job.requests), 'Capturas validadas': statuses.count('Captura validada'),
            'Descargadas': result.downloaded if result else 0, 'Sin cambios': result.unchanged if result else 0,
            'Omitidas por caché': result.skipped_existing if result else 0,
            'No disponibles': result.unavailable if result else 0, 'Fallidas': result.failed if result else 0,
            'Inicio UTC': started, 'Fin UTC': utc_now_iso(), 'Error': error})
    exit_code = 1 if any(j['Estado'] == 'Falló' for j in jobs) else 2 if any(j['Estado'] == 'Pendiente' for j in jobs) else 0
    return jobs, coverage, exit_code
