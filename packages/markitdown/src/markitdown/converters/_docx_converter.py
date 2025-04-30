import sys
import os
from typing import BinaryIO, Any

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

    def _get_document_name(self, stream_info: StreamInfo) -> str:
        """
        Extract document name from StreamInfo
        """
        # First try to extract from filename attribute
        if stream_info.filename:
            basename = os.path.basename(stream_info.filename)
            name, _ = os.path.splitext(basename)
            if name:
                print(f"[DEBUG] Extracted document name from filename: {name}")
                return name
        
        # If local_path exists, try to extract from local path
        if stream_info.local_path:
            basename = os.path.basename(stream_info.local_path)
            name, _ = os.path.splitext(basename)
            if name:
                print(f"[DEBUG] Extracted document name from local_path: {name}")
                return name
                
        # If URL exists, try to extract from URL
        if stream_info.url:
            basename = os.path.basename(stream_info.url)
            name, _ = os.path.splitext(basename)
            if name:
                print(f"[DEBUG] Extracted document name from URL: {name}")
                return name
        
        # Default name
        return "docx_document"

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        print(f"[DEBUG] DocxConverter.convert called with kwargs: {kwargs}")
        print(f"[DEBUG] StreamInfo: filename={stream_info.filename}, local_path={stream_info.local_path}, url={stream_info.url}")
        
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
            print(f"[DEBUG] Setting conversion_name to: {conversion_name}")

        style_map = kwargs.get("style_map", None)
        pre_process_stream = pre_process_docx(file_stream)
        
        # Convert to HTML and pass necessary parameters to HTML converter
        html_content = mammoth.convert_to_html(pre_process_stream, style_map=style_map).value
        
        # Create new StreamInfo to pass to HTML converter
        html_stream_info = stream_info.copy_and_update(
            mimetype="text/html",
            extension=".html"
        )
        
        print(f"[DEBUG] Calling HTML converter with parameters: conversion_name={kwargs.get('conversion_name')}")
        # Use io.BytesIO to create binary stream
        from io import BytesIO
        return self._html_converter.convert(
            file_stream=BytesIO(html_content.encode("utf-8")),
            stream_info=html_stream_info,
            **kwargs,
        )