from xml.dom.minidom import *
from zipfile import ZipFile
from datetime import datetime


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

    # Helper functions
    def createElementWithText(self, tagName, text=None):
        el = self.xml.createElement(tagName)
        if text is not None:
            el.appendChild(self.xml.createTextNode(text))
        return el

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

        aux["[Content_Types].xml"] = ctxml

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
        aux["_rels/.rels"] = rels

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
        aux["word/_rels/document.xml.rels"] = word_rels

        # Document properties

        # Core file
        cur_time = datetime.utcnow().strftime("%Y-%m-%dT%XZ")
        coxml = Document()
        corexml = self.xml.createElement("cp:coreProperties")
        core_attrs = (
            ("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"),
            ("dc", "http://purl.org/dc/elements/1.1/"),
            ("dcmitype", "http://purl.org/dc/dcmitype/"),
            ("dcterms", "http://purl.org/dc/terms/"),
            ("xsi", "http://www.w3.org/2001/XMLSchema-instance")
        )
        for name, value in core_attrs:
            corexml.setAttribute("xmlns:%s" % name, value)
        # create nodes from properties
        nodes = (
            ("dc:title", "title"),
            ("dc:subject", "subject"),
            ("dc:creator", "creator"),
            ("cp:lastModifiedBy", "creator"),
            ("dc:description", "description"),
            ("dc:keywords", "keywords"),
            ("cp:revision", "revision"),
        )
        for name, prop in nodes:
            el = self.createElementWithText(name, self.getProperty(prop))
            corexml.appendChild(el)

        # created element
        created = self.createElementWithText("dcterms:created", cur_time)
        created.setAttribute("xsi:type","dcterms:W3CDTF")
        corexml.appendChild(created)
        # modified element
        modified = self.createElementWithText("dcterms:modified", cur_time)
        modified.setAttribute("xsi:type","dcterms:W3CDTF")
        corexml.appendChild(modified)

        coxml.appendChild(corexml)

        aux["docProps/core.xml"] = coxml

        # Apps file
        appxml = parseString('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" />')
        # create fake info for infoNodes
        # TODO: Actually calculate this
        infoNodes = ("TotalTime", "Pages", "Words", "Characters", "Lines", "Paragraphs", "CharactersWithSpaces")
        for name in infoNodes:
            el = self.createElementWithText(name, "1")
            appxml.documentElement.appendChild(el)
        # Template information, etc.
        otherNodes = (
            ("Template", "Normal.dotm"),
            ("Application", "Microsoft Office Word"),
            ("DocSecurity", "0"),
            ("ScaleCrop", "false"),
            ("Company", None),
            ("LinksUpToDate", "false"),
            ("SharedDoc", "false"),
            ("HyperlinksChanged", "false"),
            ("AppVersion", "15.0000")
        )
        for name, value in otherNodes:
            el = self.createElementWithText(name, value)
            appxml.documentElement.appendChild(el)
        # Nested nodes (HeadingPairs, TitlesOfParts)
        hp = appxml.createElement("HeadingPairs")
        hpvec = appxml.createElement("vt:vector")
        hpvec.setAttribute("baseType", "variant")
        hpvec.setAttribute("size", "2")
        var1 = appxml.createElement("vt:variant")
        var1.appendChild(self.createElementWithText("vt:lpstr", "Title"))
        hpvec.appendChild(var1)
        var2 = appxml.createElement("vt:variant")
        var2.appendChild(self.createElementWithText("vt:i4", "1"))
        hpvec.appendChild(var2)
        hp.appendChild(hpvec)
        tp = appxml.createElement("TitlesOfParts")
        tpvec = appxml.createElement("vt:vector")
        tpvec.setAttribute("baseType", "lpstr")
        tpvec.setAttribute("size", "1")
        tpvec.appendChild(appxml.createElement("vt:lpstr"))
        tp.appendChild(tpvec)
        appxml.documentElement.appendChild(hp)
        appxml.documentElement.appendChild(tp)

        aux["docProps/app.xml"] = appxml

        # Font Table

        # Settings

        # Styles

        # Web Settings

        return aux

    def save(self, fn):
        aux = self.makeAuxFiles()
        z = ZipFile(fn, "w")
        # write auxiliary files
        for fn, content in aux.items():
            z.writestr(fn, content.toprettyxml(encoding="UTF-8")[:-1])
        # write document to zip file
        z.writestr("word/document.xml","%s\n%s" % ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',self.document.toprettyxml(encoding="UTF-8")[:-1]))
        z.close()
