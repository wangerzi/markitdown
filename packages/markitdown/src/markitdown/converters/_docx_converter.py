import sys
import os
import re
import unicodedata
from typing import BinaryIO, Any
from io import BytesIO

from ._html_converter import HtmlConverter
from ..converter_utils.docx.pre_process import pre_process_docx
from .._base_converter import DocumentConverter, DocumentConverterResult
from .._stream_info import StreamInfo
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE

# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_dependency_exc_info = None
try:
    import mammoth
except ImportError:
    # Preserve the error and stack trace for later
    _dependency_exc_info = sys.exc_info()


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

ACCEPTED_FILE_EXTENSIONS = [".docx"]


class DocxConverter(HtmlConverter):
    """
    Converts DOCX files to Markdown. Style information (e.g.m headings) and tables are preserved where possible.
    """

    def __init__(self):
        super().__init__()
        self._html_converter = HtmlConverter()

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False
    
    def _sanitize_filename(self, filename: str) -> str:
        """
        Sanitize a filename by removing or replacing problematic characters.
        
        Args:
            filename: The original filename
            
        Returns:
            A sanitized filename safe for filesystem use
        """
        # Step 1: Normalize unicode characters
        filename = unicodedata.normalize('NFKD', filename)
        
        # Step 2: Remove invalid characters and replace spaces with underscores
        # Keep alphanumeric characters, underscores, hyphens, and periods
        sanitized = re.sub(r'[^\w\-\.]', '_', filename)
        
        # Step 3: Collapse multiple underscores
        sanitized = re.sub(r'_+', '_', sanitized)
        
        # Step 4: Remove leading/trailing underscores
        sanitized = sanitized.strip('_')
        
        # Step 5: Ensure we have a valid filename (default if empty)
        if not sanitized:
            sanitized = "unnamed"
        
        return sanitized

    def _get_document_name(self, stream_info: StreamInfo) -> str:
        """
        Extract document name from StreamInfo
        """
        # First try to extract from filename attribute
        if stream_info.filename:
            basename = os.path.basename(stream_info.filename)
            name, _ = os.path.splitext(basename)
            if name:
                return self._sanitize_filename(name)
        
        # If local_path exists, try to extract from local path
        if stream_info.local_path:
            basename = os.path.basename(stream_info.local_path)
            name, _ = os.path.splitext(basename)
            if name:
                return self._sanitize_filename(name)
                
        # If URL exists, try to extract from URL
        if stream_info.url:
            basename = os.path.basename(stream_info.url)
            name, _ = os.path.splitext(basename)
            if name:
                return self._sanitize_filename(name)
        
        # Default name
        return "docx_document"

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:     
        # Check dependencies
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".docx",
                    feature="docx",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        # If conversion_name not explicitly provided, try to extract from stream_info
        if "conversion_name" not in kwargs:
            conversion_name = self._get_document_name(stream_info)
            kwargs["conversion_name"] = conversion_name

        style_map = kwargs.get("style_map", None)
        pre_process_stream = pre_process_docx(file_stream)
        
        # Convert to HTML and pass necessary parameters to HTML converter
        html_content = mammoth.convert_to_html(pre_process_stream, style_map=style_map).value
        
        # Create new StreamInfo to pass to HTML converter
        html_stream_info = stream_info.copy_and_update(
            mimetype="text/html",
            extension=".html"
        )
        
        # Use io.BytesIO to create binary stream
        from io import BytesIO
        return self._html_converter.convert(
            file_stream=BytesIO(html_content.encode("utf-8")),
            stream_info=html_stream_info,
            **kwargs,
        )