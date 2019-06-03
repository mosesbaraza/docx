#by Baraza Moses
#email:<clonesoftwares@gmail.com>
#This piece of software is under Python free sotware License

class DocxFile:
    def __init__(self, docx_name):
        self.docx_name=docx_name
        self.expandDocxFile() #method to expand word document

    def expandDocxFile(self):
        import zipfile
        try:
            with zipfile.ZipFile(self.docx_name, 'r') as zipStream:
                with zipStream.open('docProps/core.xml') as fileStream:
                    import xml.etree.ElementTree as ET
                    root=ET.fromstring(fileStream.read())
                    for child in root:
                        if child.tag=='{http://purl.org/dc/elements/1.1/}creator':
                            self.creator_name=child.text
                        elif child.tag=='{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy':
                            self.file_modified_by=child.text
                        elif child.tag=='{http://purl.org/dc/terms/}created':
                            self.file_creation_date=child.text
                        elif child.tag=='{http://purl.org/dc/terms/}modified':
                            self.file_modified_date=child.text
                fileStream.close()

                self.text_list=[] #to carry strings from the document
                with zipStream.open('word/document.xml', 'r') as docstream: ##read document.xml since it contains the text in the document
                    root2=ET.fromstring(docstream.read())
                    for child in root2:
                        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body':
                            for subchild in child:
                                if subchild.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
                                    for subsubchild in subchild:
                                        if subsubchild.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r':
                                            for x in subsubchild:
                                                self.text_list.append(x.text)
                
                docstream.close()
            zipStream.close()
        except zipfile.BadZipFile:
            pass

    def createdby(self):
        return self.creator_name

    def modifiedby(self):
        return self.file_modified_by

    def creationdatetime(self):
        return self.file_creation_date
    
    def modifydatetime(self):
        return self.file_modified_date

    def readdocx(self):
        return self.text_list

