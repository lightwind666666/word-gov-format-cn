# Adding Red Separator Line Under Issuer Mark

The formatter does NOT add the red horizontal line beneath the issuer mark (发文机关标识). This is a required visual element in Chinese official documents. Add it manually after formatting.

## Code

```python
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document('file.docx')

# First paragraph should be the issuer mark (红色抬头)
header_p = doc.paragraphs[0]
pPr = header_p._element.get_or_add_pPr()
pBdr = OxmlElement('w:pBdr')

bottom = OxmlElement('w:bottom')
bottom.set(qn('w:val'), 'single')    # single line
bottom.set(qn('w:sz'), '12')         # ~1pt
bottom.set(qn('w:space'), '4')       # gap between text and line
bottom.set(qn('w:color'), 'FF0000')  # red

pBdr.append(bottom)
pPr.append(pBdr)

doc.save('file.docx')
```
