import os
import tempfile
import zipfile
from xml.etree import ElementTree as ET

class DocxFile:
    def __init__(self, file):
        self.file=file
        self.docxextractor()

    def docxextractor(self):
        self.temp_dir=tempfile.TemporaryDirectory()
        os.chdir(self.temp_dir.name)
        with zipfile.ZipFile(self.file, 'r') as zp:
            zp.extractall(".")
        zp.close()

    def changecreator(self, name):
        xmlstream=ET.parse('docProps/core.xml')
        for element in xmlstream.iter():
            if element.tag=="{http://purl.org/dc/elements/1.1/}creator":
                element.text=name
            else:
                pass
        os.remove('docProps/core.xml')
        os.remove(self.file)
        xmlstream.write('docProps/core.xml')

        with zipfile.ZipFile(self.file, 'w') as strm:
            for data in os.walk("."):
                for direntry in os.scandir(data[0]):
                    strm.write(direntry.path)

        strm.close()

x=DocxFile('/home/moses/Documents/doc/seth.docx')
x.changecreator('Maina')

    
