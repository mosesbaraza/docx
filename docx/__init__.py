from xml.etree import ElementTree as ET


class DocxFile: #class that will create an object to capture file to be read
    def __init__(self, filename):
        self.filename=filename
        self.docxstream

    @property
    def docxstream(self):  #attribute that provides the filestream of the docx file
        try:
            import zipfile
        except ImportError:
            raise ImportError('Failed to import zipfile')
        else:
            docx_stream = zipfile.ZipFile(self.filename)
            return docx_stream

    def read(self): #method to read the content of docx file and returns a list of lines within the file
        txt_list=[]
        xmlstream=ET.fromstring(self.docxstream.read('word/document.xml').decode(encoding='utf-8')) #word/documents.xml contains the content of the .docx file
        for element in xmlstream.iter():
            if element.text!=None:
                txt_list.append(element.text)
            else:
                pass

        return txt_list

    def creator(self): #get the original creator of the document
        xmlstream=ET.fromstring(self.docxstream.read('docProps/core.xml').decode(encoding='utf-8'))
        for element in xmlstream.iter():
            if element.tag=='{http://purl.org/dc/elements/1.1/}creator':
                return element.text
            else:
                pass

    def lastmodifiedby(self): #get the user to last modify the document
        xmlstream=ET.fromstring(self.docxstream.read('docProps/core.xml').decode(encoding='utf-8'))
        for element in xmlstream.iter():
            if element.tag=='{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy':
                return element.text
            else:
                pass

    def datetimecreated(self): #return the datetime of document creation
        xmlstream=ET.fromstring(self.docxstream.read('docProps/core.xml').decode(encoding='utf-8'))
        for element in xmlstream.iter():
            if element.tag=='{http://purl.org/dc/terms/}created':
                return element.text
            else:
                pass

    def datetimemodified(self): #return the datetime of file modification
        xmlstream=ET.fromstring(self.docxstream.read('docProps/core.xml').decode(encoding='utf-8'))
        for element in xmlstream.iter():
            if element.tag=='{http://purl.org/dc/terms/}modified':
                return element.text
            else:
                pass

    def docxstat(self): #return a dictionary of the document statistics
        stat = dict(zip(['creator', 'last_modified_by', 'datetime_created', 'datetime_modified'], [self.creator(), self.lastmodifiedby(), self.datetimecreated(), self.datetimemodified()]))
        return stat


#path_='/home/moses/Documents/doc/seth.docx'
#x=DocxFile(path_)
#print(x.creator())
#print(x.lastmodifiedby())
#print(x.datetimecreated())
#print(x.datetimemodified())
#print(x.docxstat())
