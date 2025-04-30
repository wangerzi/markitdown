import re
import markdownify
import os
import base64
import hashlib
import sys

from typing import Any, Optional
from urllib.parse import quote, unquote, urlparse, urlunparse


class _CustomMarkdownify(markdownify.MarkdownConverter):
    """
    A custom version of markdownify's MarkdownConverter. Changes include:

    - Altering the default heading style to use '#', '##', etc.
    - Removing javascript hyperlinks.
    - Truncating images with large data:uri sources.
    - Ensuring URIs are properly escaped, and do not conflict with Markdown syntax
    """

    def __init__(self, **options: Any):
        # Set default values for image-related options
        self.image_output_dir = options.get("image_output_dir", "assets")
        self.conversion_name = options.get("conversion_name")
        
        # Apply basic options
        options["heading_style"] = options.get("heading_style", markdownify.ATX)
        options["keep_data_uris"] = options.get("keep_data_uris", False)
        
        # Initialize parent class
        super().__init__(**options)

    def convert_hn(
        self,
        n: int,
        el: Any,
        text: str,
        convert_as_inline: Optional[bool] = False,
        **kwargs,
    ) -> str:
        """Same as usual, but be sure to start with a new line"""
        if not convert_as_inline:
            if not re.search(r"^\n", text):
                return "\n" + super().convert_hn(n, el, text, convert_as_inline)  # type: ignore

        return super().convert_hn(n, el, text, convert_as_inline)  # type: ignore

    def convert_a(
        self,
        el: Any,
        text: str,
        convert_as_inline: Optional[bool] = False,
        **kwargs,
    ):
        """Same as usual converter, but removes Javascript links and escapes URIs."""
        prefix, suffix, text = markdownify.chomp(text)  # type: ignore
        if not text:
            return ""

        if el.find_parent("pre") is not None:
            return text

        href = el.get("href")
        title = el.get("title")

        # Escape URIs and skip non-http or file schemes
        if href:
            try:
                parsed_url = urlparse(href)  # type: ignore
                if parsed_url.scheme and parsed_url.scheme.lower() not in ["http", "https", "file"]:  # type: ignore
                    return "%s%s%s" % (prefix, text, suffix)
                href = urlunparse(parsed_url._replace(path=quote(unquote(parsed_url.path))))  # type: ignore
            except ValueError:  # It's not clear if this ever gets thrown
                return "%s%s%s" % (prefix, text, suffix)

        # For the replacement see #29: text nodes underscores are escaped
        if (
            self.options["autolinks"]
            and text.replace(r"\_", "_") == href
            and not title
            and not self.options["default_title"]
        ):
            # Shortcut syntax
            return "<%s>" % href
        if self.options["default_title"] and not title:
            title = href
        title_part = ' "%s"' % title.replace('"', r"\"") if title else ""
        return (
            "%s[%s](%s%s)%s" % (prefix, text, href, title_part, suffix)
            if href
            else text
        )

    def convert_img(
        self,
        el: Any,
        text: str,
        convert_as_inline: Optional[bool] = False,
        **kwargs,
    ) -> str:
        """
        Process image elements, save data URI format images to filesystem
        Supports categorized storage in subfolders by document name
        """
        alt = el.attrs.get("alt", None) or ""
        src = el.attrs.get("src", None) or el.attrs.get("data-src", None) or ""
        title = el.attrs.get("title", None) or ""
        title_part = ' "%s"' % title.replace('"', r"\"") if title else ""
        
        # If in inline mode and not preserved, return alt text
        if (
            convert_as_inline
            and el.parent.name not in self.options.get("keep_inline_images_in", [])
        ):
            return alt

        # Process data URI format images
        if src.startswith("data:image") and not self.options.get("keep_data_uris", False):
            try:
                # Parse MIME type
                mime_type = src.split(";")[0].replace("data:", "")
                
                # Get file extension
                ext = {
                    "image/png": ".png",
                    "image/jpeg": ".jpg",
                    "image/jpg": ".jpg",
                    "image/gif": ".gif"
                }.get(mime_type, ".png")
                
                # Decode base64 data
                encoded = src.split(",")[1]
                image_data = base64.b64decode(encoded)
                
                # Generate unique filename
                hashname = hashlib.sha256(image_data).hexdigest()[:8]
                filename = f"image_{hashname}{ext}"
                
                # Determine output directory
                if hasattr(self, 'conversion_name') and self.conversion_name:
                    # If conversion_name exists, create subfolder
                    output_dir = os.path.join(self.image_output_dir, self.conversion_name)
                else:
                    # Otherwise use base directory
                    output_dir = self.image_output_dir
                
                # Ensure directory exists
                os.makedirs(output_dir, exist_ok=True)
                
                # Save image file
                filepath = os.path.join(output_dir, filename)
                with open(filepath, "wb") as f:
                    f.write(image_data)
                
                # Update src to relative path
                src = os.path.join(output_dir, filename).replace("\\", "/")
                
            except Exception as e:
                error_msg = f"Error saving image: {str(e)}"
                import traceback
                traceback.print_exc(file=sys.stderr)
                # If extraction fails, revert to original truncating behavior
                src = src.split(",")[0] + "..."
                return f"![{alt}](image_error.png)  <!-- {error_msg} -->"

        # Process other data URIs that are not images (truncate them)
        elif src.startswith("data:") and not self.options.get("keep_data_uris", False):
            src = src.split(",")[0] + "..."

        # Return Markdown format image reference
        return f"![{alt}]({src}{title_part})"

    def convert_soup(self, soup: Any) -> str:
        return super().convert_soup(soup)  # type: ignore
