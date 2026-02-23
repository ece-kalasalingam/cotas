class ValidationError(Exception):
    """Raised when strict validation fails."""
    pass

class SystemError(Exception):
    """Raised when unexpected internal failure occurs."""
    pass