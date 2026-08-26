# -*- coding: utf-8 -*-
"""Build a tiny test .docx containing flattened Cambria Math formulas."""
import zipfile, shutil, os, uuid

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

doc = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="%s" xmlns:m="%s">
<w:body>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>Err = x</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>0</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t xml:space="preserve"> + x</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>i</w:t></w:r>
</w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>s = &#8730;[&#931;</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>&#7536;&#7528;&#8320;</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t> (x</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>i</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t xml:space="preserve"> &#8722; x</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>&#772;)</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>&#178; / (n &#8722; 1)]</w:t></w:r>
</w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>P(k) = &#955;</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>&#7486;</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t xml:space="preserve"> e</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t>^{&#8722;&#955;}</w:t></w:r>
<w:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><w:t xml:space="preserve"> / k!</w:t></w:r>
</w:p>
<w:sectPr/></w:body>
</w:document>
''' % (W, _M)

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

out = r'C:\Users\vipuser\Desktop\formula-repair\_sample_in.docx'
if os.path.exists(out):
    os.remove(out)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', root_rels)
    z.writestr('word/document.xml', doc)
print('wrote', out, os.path.getsize(out))