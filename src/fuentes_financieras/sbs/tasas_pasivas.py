from __future__ import annotations

from fuentes_financieras.api import source

_DATASET_ID = "pe.sbs.tasas_pasivas"


def fetch(**query):
    return source(_DATASET_ID).fetch(**query)


def sync(**query):
    return source(_DATASET_ID).sync(**query)


def load(**query):
    return source(_DATASET_ID).load(**query)


def provider():
    return source(_DATASET_ID)
