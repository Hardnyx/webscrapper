from __future__ import annotations

from fuentes_financieras.api import source


DATASET_ID = "pe.sbs.clasificaciones_riesgo"


def _provider():
    return source(DATASET_ID)


def fetch(**query):
    return _provider().fetch(**query)


def sync(**query):
    return _provider().sync(**query)


def load(**query):
    return _provider().load(**query)


def available_periods():
    return _provider().available_periods()
