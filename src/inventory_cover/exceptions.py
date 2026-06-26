"""Custom exceptions for the inventory cover engine."""

from __future__ import annotations

from dataclasses import dataclass, field


class PipelineError(Exception):
    """Base class for pipeline errors."""


class CatastrophicPipelineError(PipelineError):
    """An error that must stop the run."""


class OutputWriteError(CatastrophicPipelineError):
    """Raised when the final output cannot be written."""


@dataclass
class FileValidationError(PipelineError):
    """A file-level validation error that should skip one file."""

    message: str
    missing_headers: list[str] = field(default_factory=list)
    details: str = ""

    def __str__(self) -> str:
        if self.missing_headers:
            return f"{self.message}: {', '.join(self.missing_headers)}"
        if self.details:
            return f"{self.message}: {self.details}"
        return self.message
