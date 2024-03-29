from xml.dom.minidom import *
from zipfile import ZipFile
from auxfiles import makeAuxFiles

class DOCXException(Exception):
    pass

class Style:
    pass

class Paragraph:
    def __init__(self, xml, text=None, style=None):
        self.xml = xml
        self.paragraph = self.xml.createElement("w:p")
        self.textRun = self.xml.createElement("w:r")
        self.paragraph.appendChild(self.textRun)
        if text is not None:
            self.setText(text)
        self.style = style

    def setText(self, text):
        t = self.xml.createElement("w:t")
        textNode = self.xml.createTextNode(text)
        t.appendChild(textNode)
        self.textRun.appendChild(t)

    def toNode(self):
        return self.paragraph

class DOCX:

    def __init__(self):
        self.xml = Document()
        # add root document node
        self.document = self.xml.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:document")
        # add namespaces
        namespaces = {
            "m":"http://schemas.openxmlformats.org/officeDocument/2006/math",
            "mc":"http://schemas.openxmlformats.org/markup-compatibility/2006",
            "o":"urn:schemas-microsoft-com:office:office",
            "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "v":"urn:schemas-microsoft-com:vml",
            "w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "w10":"urn:schemas-microsoft-com:office:word",
            "w14":"http://schemas.microsoft.com/office/word/2010/wordml",
            "w15":"http://schemas.microsoft.com/office/word/2012/wordml",
            "wne":"http://schemas.microsoft.com/office/word/2006/wordml",
            "wp":"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "wp14":"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "wpc":"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            "wpg":"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            "wpi":"http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            "wps":"http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        }
        self.document.setAttribute("mc:Ignorable", "w14 w15 wp14")
        for name, uri in namespaces.items():
            self.document.setAttribute("xmlns:%s" % name, uri)
        # add body node
        self.body = self.xml.createElement("w:body")
        self.document.appendChild(self.body)
        self.properties = {"title": None, "subject": None, "creator": None, "keywords": None, "description": None, "revision": "1"}

    # Properties
    def setProperty(self, prop, value):
        if prop not in self.properties.keys():
            raise DOCXException("%s is not a valid DOCX property" % prop)
        self.properties[prop] = value

    def getProperty(self, prop):
        if prop not in self.properties.keys():
            raise DOCXException("%s is not a valid DOCX property" % prop)
        # just return keywords as a string
        if prop == "keywords" and self.properties[prop] is not None:
            return ", ".join(self.properties[prop])
        return self.properties[prop]

    # Elements
    def paragraph(self, text=None, style=None):
        return Paragraph(self.xml, text, style)

    def add(self, element):
        self.body.appendChild(element.toNode())

    def save(self, fn):
        aux = makeAuxFiles(self)
        z = ZipFile(fn, "w")
        # write auxiliary files
        for fn, content in aux.items():
            z.writestr(fn, content.toprettyxml(encoding="UTF-8")[:-1])
        # write document to zip file
        z.writestr("word/document.xml","%s\n%s" % ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',self.document.toprettyxml(encoding="UTF-8")[:-1]))
        z.close()
