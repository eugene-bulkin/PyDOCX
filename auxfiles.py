from xml.dom.minidom import *
from datetime import datetime

def createElementWithText(tagName, text=None):
    xml = Document()
    el = xml.createElement(tagName)
    if text is not None:
        el.appendChild(xml.createTextNode(text))
    return el

def contentTypes():
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

    return ctxml

def relationshipFiles():
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

    return (rels, word_rels)

def coreXML(docx):
    cur_time = datetime.utcnow().strftime("%Y-%m-%dT%XZ")
    coxml = Document()
    corexml = coxml.createElement("cp:coreProperties")
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
        el = createElementWithText(name, docx.getProperty(prop))
        corexml.appendChild(el)

    # created element
    created = createElementWithText("dcterms:created", cur_time)
    created.setAttribute("xsi:type","dcterms:W3CDTF")
    corexml.appendChild(created)
    # modified element
    modified = createElementWithText("dcterms:modified", cur_time)
    modified.setAttribute("xsi:type","dcterms:W3CDTF")
    corexml.appendChild(modified)

    coxml.appendChild(corexml)

    return coxml

def appXML():
    appxml = parseString('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" />')
    # create fake info for infoNodes
    # TODO: Actually calculate this
    infoNodes = ("TotalTime", "Pages", "Words", "Characters", "Lines", "Paragraphs", "CharactersWithSpaces")
    for name in infoNodes:
        el = createElementWithText(name, "1")
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
        el = createElementWithText(name, value)
        appxml.documentElement.appendChild(el)
    # Nested nodes (HeadingPairs, TitlesOfParts)
    hp = appxml.createElement("HeadingPairs")
    hpvec = appxml.createElement("vt:vector")
    hpvec.setAttribute("baseType", "variant")
    hpvec.setAttribute("size", "2")
    var1 = appxml.createElement("vt:variant")
    var1.appendChild(createElementWithText("vt:lpstr", "Title"))
    hpvec.appendChild(var1)
    var2 = appxml.createElement("vt:variant")
    var2.appendChild(createElementWithText("vt:i4", "1"))
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

    return appxml

def makeAuxFiles(docx):
    aux = {}

    aux["[Content_Types].xml"] = contentTypes()
    aux["_rels/.rels"], aux["word/_rels/document.xml.rels"] = relationshipFiles()
    aux["docProps/core.xml"] = coreXML(docx)
    aux["docProps/app.xml"] = appXML()

    return aux