class PipelineError(RuntimeError):
    def __init__(
        self, message: str, status_code: int = 500, *, public_message: str | None = None,
    ):
        super().__init__(message)
        self.status_code = status_code
        # Only fixed, application-owned messages may be marked public.
        self.public_message = public_message


def get_public_error_message(error: Exception) -> str:
    """Keep backend exception details out of public responses and job records."""
    if isinstance(error, PipelineError) and error.public_message:
        return error.public_message
    return "Unable to process this document. Please try again later."


def get_error_status_code(error: Exception, default: int = 500) -> int:
    status_code = getattr(error, "status_code", None)
    if isinstance(status_code, int):
        return status_code
    return default


def format_error_message(error: Exception, fallback: str = "Unknown error") -> str:
    message = str(error).strip()
    if message:
        return message
    return fallback
