# puralox/base_importer.py
from abc import ABC, abstractmethod
from typing import Optional


class BaseImporter(ABC):
    """
    Abstract base class that defines the common import contract for all file importers.
    Concrete implementations: ExcelProcessor, PdfProcessor.
    """

    @abstractmethod
    def import_file(self, path: str, original_filename: Optional[str] = None) -> int:
        """Import a file into the database and return the new file_info_id."""
        raise NotImplementedError
