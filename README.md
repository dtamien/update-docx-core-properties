# Update DOCX Core Properties

In the context of developing a conversational assistant, I needed to export its responses into a `.docx` template. One requirement was that the document's author and modification dates should reflect the assistant's identity and the time of generation.

In the [update_docx_core_properties.py](update_docx_core_properties.py) file, I propose a straightforward Python function to overwrite the core properties of a DOCX file.

## DOCX Core Properties

DOCX files are essentially ZIP archives containing XML files. The core properties of a DOCX document are stored in the `docProps/core.xml` file within the archive. These properties include:

- Author
- Category
- Comments
- Content Status
- Created Date
- Identifier
- Keywords
- Language
- Last Modified By
- Last Printed
- Modified Date
- Revision
- Subject
- Title
- Version

## Display Core Properties

### From Windows

They can be viewed by right-clicking the file and selecting `Properties` and then `Details`.

### From macOS and Linux

Using `unzip` and `xmllint`, you can inspect them with the following command:

```sh
unzip -p sample.docx docProps/core.xml | xmllint --format -
```

A typical output looks like this:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Word Sample</dc:title>
  <dc:subject/>
  <dc:creator>dtamien</dc:creator>
  <cp:keywords/>
  <dc:description/>
  <cp:lastModifiedBy>dtamien</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2025-10-03T18:46:35Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2025-10-03T18:46:35Z</dcterms:modified>
</cp:coreProperties>
```
