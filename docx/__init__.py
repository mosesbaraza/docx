import zipfile
import os
import xml.etree.ElementTree as ET



class DocxFile:
    def __init__(self, docx_name):
        self.docx_name=docx_name
        self.creator_name=''
        self.file_modified_by=''
        self.file_creation_date=''
        self.file_modified_date=''
        self.expandDocxFile()

    def expandDocxFile(self):
            try:
                with zipfile.ZipFile(self.docx_name, 'r') as zipStream:
                    with zipStream.open('docProps/core.xml') as fileStream:
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
                zipStream.close()
            except zipfile.BadZipFile:
                pass

    def fileCreatedBy(self):
        return self.creator_name

    def fileModifiedBy(self):
        return self.file_modified_by

    def fileCreationDatetime(self):
        return self.file_creation_date
    
    def fileModifiedDatetime(self):
        return self.file_modified_date
