class SMVError(RuntimeError):
    """Error base de la fuente SMV."""


class SMVRequestError(SMVError):
    """La página de SMV no pudo recuperarse correctamente."""


class SMVParseError(SMVError):
    """La respuesta de SMV no contiene la estructura esperada."""
