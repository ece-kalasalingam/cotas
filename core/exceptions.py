class ValidationError(Exception):
    """
    Raised when strict validation fails.
    Intended for user-facing validation errors.
    """
    pass


class SystemError(Exception):
    """
    Raised for unexpected internal failures.
    Indicates a logic or system-level issue.
    """
    pass