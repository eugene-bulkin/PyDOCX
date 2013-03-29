from xml.dom.minidom import *
from zipfile import ZipFile

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

    # Elements
    def paragraph(self, text=None, style=None):
        return Paragraph(self.xml, text, style)

    def add(self, element):
        self.body.appendChild(element.toNode())

    def makeAuxFiles(self):
        aux = {}

        # Content Types file
        ctxml = parseString('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" />')
        nodes = (
                    ("Default", "application/vnd.openxmlformats-package.relationships+xml", "rels"),
                    ("Default", "application/xml", "xml"),
                    ("Override","application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "/word/document.xml"),
                    ("Override","application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", "/word/styles.xml"),
                    ("Override","application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", "/word/settings.xml"),
                    ("Override","application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml", "/word/webSettings.xml"),
                    ("Override","application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml", "/word/fontTable.xml"),
                    ("Override","application/vnd.openxmlformats-officedocument.theme+xml", "/word/theme/theme1.xml"),
                    ("Override","application/vnd.openxmlformats-package.core-properties+xml", "/docProps/core.xml"),
                    ("Override","application/vnd.openxmlformats-officedocument.extended-properties+xml", "/docProps/app.xml")
        )

        for tag, a1, a2 in nodes:
            el = ctxml.createElement(tag)
            el.setAttribute("ContentType", a1) # first attribute is always ContentType
            # set second attribute
            if tag == "Default":
                el.setAttribute("Extension", a2)
            else:
                el.setAttribute("PartName", a2)
            ctxml.documentElement.appendChild(el)

        aux["[Content_Types].xml"] = ctxml.toprettyxml()

        # Relationship files
        rels = parseString('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />')
        rels_attrs = (
                ("rId3", "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
                ("rId2", "docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"),
                ("rId1", "word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        )
        for i, targ, ty in rels_attrs:
            el = rels.createElement("Relationship")
            el.setAttribute("Id", i)
            el.setAttribute("Target", targ)
            el.setAttribute("Type", ty)
            rels.documentElement.appendChild(el)
        aux["_rels/.rels"] = rels.toprettyxml()

        word_rels = parseString('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />')
        word_rels_attrs = (
                ("rId3", "webSettings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"),
                ("rId2", "settings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"),
                ("rId1", "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
                ("rId5", "theme/theme1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                ("rId4", "fontTable.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable")
        )
        for i, targ, ty in word_rels_attrs:
            el = word_rels.createElement("Relationship")
            el.setAttribute("Id", i)
            el.setAttribute("Target", targ)
            el.setAttribute("Type", ty)
            rels.documentElement.appendChild(el)
        aux["word/_rels/document.xml.rels"] = word_rels.toprettyxml()

        return aux

    def save(self, fn):
        aux = self.makeAuxFiles()
        z = ZipFile(fn, "w")
        # write auxiliary files
        for fn, content in aux.items():
            z.writestr(fn, content)
        # write document to zip file
        z.writestr("word/document.xml","%s\n%s" % ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',self.document.toprettyxml(encoding="UTF-8")[:-1]))
        z.close()
