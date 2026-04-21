"""
Custom exception hierarchy for the RPA tool.
All exceptions are subclasses of RPAError for easy catch-all handling.
"""


class RPAError(Exception):
    """Base class for all RPA tool errors."""


# --- Recording errors ---

class RecorderError(RPAError):
    """Raised when the recorder encounters an unrecoverable error."""


class HookInstallError(RecorderError):
    """Failed to install a system-level input hook."""


class ElevationRequiredError(RecorderError):
    """Target process requires administrator elevation to be captured."""


# --- Replay errors ---

class ReplayError(RPAError):
    """Raised when replay fails at an event."""
    def __init__(self, message: str, event_id: int | None = None):
        super().__init__(message)
        self.event_id = event_id


class ElementNotFoundError(ReplayError):
    """Could not locate the target UI element using any matching strategy."""


class ElementNotInteractableError(ReplayError):
    """Element was found but is disabled, hidden, or off-screen."""


class WindowNotFoundError(ReplayError):
    """The target application window could not be found."""


class ReplayTimeoutError(ReplayError):
    """Timed out waiting for the UI to become ready."""


# --- Storage errors ---

class StorageError(RPAError):
    """File I/O or database error."""


class SessionNotFoundError(StorageError):
    """Requested session ID does not exist."""


class SessionCorruptError(StorageError):
    """Session file exists but cannot be parsed."""


class SchemaMismatchError(StorageError):
    """Session file schema version is incompatible."""


# --- Config errors ---

class ConfigError(RPAError):
    """Invalid or missing configuration."""
