from typing import Optional
from datetime import datetime
from zoneinfo import ZoneInfo

from docx import Document
from docx.opc.coreprops import CoreProperties


def update_docx_core_properties(
    template_path: str,
    output_path: str,
    author: Optional[str] = None,
    category: Optional[str] = None,
    comments: Optional[str] = None,
    content_status: Optional[str] = None,
    created: Optional[datetime] = None,
    identifier: Optional[str] = None,
    keywords: Optional[str] = None,
    language: Optional[str] = None,
    last_modified_by: Optional[str] = None,
    last_printed: Optional[datetime] = None,
    modified: Optional[datetime] = None,
    revision: Optional[int] = None,
    subject: Optional[str] = None,
    title: Optional[str] = None,
    version: Optional[str] = None,
) -> str:
    """
    Update core properties of a Word document (.docx).
    Only replaces properties if value is not None.
    """
    document = Document(template_path)
    core_properties: CoreProperties = document.core_properties

    values = dict(
        author=author,
        category=category,
        comments=comments,
        content_status=content_status,
        created=created,
        identifier=identifier,
        keywords=keywords,
        language=language,
        last_modified_by=last_modified_by,
        last_printed=last_printed,
        modified=modified,
        revision=revision,
        subject=subject,
        title=title,
        version=version,
    )

    for key, value in values.items():
        if value is not None:
            setattr(core_properties, key, value)

    document.save(output_path)
    return output_path


if __name__ == "__main__":
    tz = ZoneInfo("Europe/Paris")
    update_docx_core_properties(
        template_path="sample.docx",
        output_path="output.docx",
        author="dtamien",
        created=datetime.now(tz),
        modified=datetime.now(tz),
        last_modified_by="dtamien",
        revision=1,
        title="Word Sample",
    )
    print("Saved output: output.docx")
