from xml.etree import ElementTree as ET
import zipfile


class DocxFile: #class that will create an object to capture file to be read
    def __init__(self, filename):
        self.filename=filename
        self.docxstream

    @property
    def docxstream(self):  #attribute that provides the filestream of the docx file
        docx_stream = zipfile.ZipFile(self.filename)
        return docx_stream

    def docxread(self): #method to read the content of docx file and returns a list of lines within the file
        txt_list=[]
        xmlstream=ET.fromstring(self.docxstream.read('word/document.xml').decode(encoding='utf-8')) #word/documents.xml contains the content of the .docx file
        for element in xmlstream.iter():
            if element.text!=None:
                txt_list.append(element.text)
            else:
                pass

        return txt_list
