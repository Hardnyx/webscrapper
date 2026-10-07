class FinancialSourcesError(RuntimeError):
    """Error base."""


class UnknownDatasetError(FinancialSourcesError, KeyError):
    """Dataset ID no registrado."""


class InvalidQueryError(FinancialSourcesError, ValueError):
    """Parámetros incompatibles con el dataset."""


class SourceUnavailableError(FinancialSourcesError):
    """La fuente externa no pudo utilizarse."""


class PeriodUnavailableError(SourceUnavailableError):
    """La fuente no publica exactamente el período solicitado."""


class WAFBlockedError(SourceUnavailableError):
    """Un WAF/CDN sustituyó la respuesta esperada."""


class SchemaChangedError(FinancialSourcesError):
    """La estructura externa cambió."""


class StorageError(FinancialSourcesError):
    """Error del almacén local."""
