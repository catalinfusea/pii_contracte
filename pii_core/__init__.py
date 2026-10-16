"""Nucleul de detecție și redactare a datelor de identificare (DOCX, PDF; .doc prin conversie Word)."""
from .engine import mask, mixed_values, norm_value, redact_file, scan_file, write_report  # noqa: F401
__version__ = "0.3.1"
