"""Public package interface for OMML to LaTeX conversion."""

from ._parser import convert_omml

__version__ = "0.1.0"
__all__ = [
    "convert_omml",
]
