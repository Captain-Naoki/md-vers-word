import sys, zipfile, pypandoc, docx
from lxml import etree
p = sys.argv[1]
z = zipfile.ZipFile(p)
for n in z.namelist():
    if n.endswith('.xml') or n.endswith('.rels'):
        etree.fromstring(z.read(n))  # lève une erreur si mal formé
print('XML bien forme :', len(z.namelist()), 'parties')
d = docx.Document(p)
print('python-docx ouvre le fichier :', len(d.paragraphs), 'paragraphes,', len(d.tables), 'tableau(x)')
out = pypandoc.convert_file(p, 'markdown', extra_args=['--wrap=none'])
sys.stdout.buffer.write(out.encode('utf-8'))
