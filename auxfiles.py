from xml.dom.minidom import *
from datetime import datetime

def createElementWithProps(tagName, text=None, attrs={}):
    xml = Document()
    el = xml.createElement(tagName)
    if text is not None:
        el.appendChild(xml.createTextNode(text))
    for attr, value in attrs.items():
        el.setAttribute(attr, value)
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
        attrs = {"Id": i, "Target": targ, "Type": ty}
        el = createElementWithProps("Relationship", attrs=attrs)
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
        el = createElementWithProps("Relationship", attrs=attrs)
        rels.documentElement.appendChild(el)

    return (rels, word_rels)

def coreXML(docx):
    cur_time = datetime.utcnow().strftime("%Y-%m-%dT%XZ")
    coxml = Document()
    core_attrs = {
        "xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "xmlns:dc": "http://purl.org/dc/elements/1.1/",
        "xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
        "xmlns:dcterms": "http://purl.org/dc/terms/",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
    }
    corexml = createElementWithProps("cp:coreProperties", attrs=core_attrs)
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
        el = createElementWithProps(name, docx.getProperty(prop))
        corexml.appendChild(el)

    # created element
    created = createElementWithProps("dcterms:created", cur_time, {"xsi:type":"dcterms:W3CDTF"})
    corexml.appendChild(created)
    # modified element
    modified = createElementWithProps("dcterms:modified", cur_time, {"xsi:type":"dcterms:W3CDTF"})
    corexml.appendChild(modified)

    coxml.appendChild(corexml)

    return coxml

def appXML():
    appxml = parseString('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" />')
    # create fake info for infoNodes
    # TODO: Actually calculate this
    infoNodes = ("TotalTime", "Pages", "Words", "Characters", "Lines", "Paragraphs", "CharactersWithSpaces")
    for name in infoNodes:
        el = createElementWithProps(name, "1")
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
        el = createElementWithProps(name, value)
        appxml.documentElement.appendChild(el)
    # Nested nodes (HeadingPairs, TitlesOfParts)
    hp = appxml.createElement("HeadingPairs")
    hpvec = createElementWithProps("vt:vector", attrs={"baseType": "variant", "size": "2"})
    var1 = appxml.createElement("vt:variant")
    var1.appendChild(createElementWithProps("vt:lpstr", "Title"))
    hpvec.appendChild(var1)
    var2 = appxml.createElement("vt:variant")
    var2.appendChild(createElementWithProps("vt:i4", "1"))
    hpvec.appendChild(var2)
    hp.appendChild(hpvec)
    tp = appxml.createElement("TitlesOfParts")
    tpvec = createElementWithProps("vt:vector", attrs={"baseType": "lpstr", "size": "1"})
    tpvec.appendChild(appxml.createElement("vt:lpstr"))
    tp.appendChild(tpvec)
    appxml.documentElement.appendChild(hp)
    appxml.documentElement.appendChild(tp)

    return appxml

def webSettings():
    wxml = Document()
    web_attrs = {
        "mc:Ignorable": "w14 w15",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml"
    }
    webxml = createElementWithProps("w:webSettings", attrs=web_attrs)
    divs = wxml.createElement("w:divs")
    div = wxml.createElement("w:div")
    div.appendChild(createElementWithProps("w:bodyDiv",attrs={"w:val":"1"}))
    div.appendChild(createElementWithProps("w:marLeft",attrs={"w:val":"0"}))
    div.appendChild(createElementWithProps("w:marRight",attrs={"w:val":"0"}))
    div.appendChild(createElementWithProps("w:marTop",attrs={"w:val":"0"}))
    div.appendChild(createElementWithProps("w:marBottom",attrs={"w:val":"0"}))
    divBdr = wxml.createElement("w:divBdr")
    bdrAttrs = {"w:color":"auto", "w:space": "0", "w:sz": "0", "w:val": "none"}
    divBdr.appendChild(createElementWithProps("w:top",attrs=bdrAttrs))
    divBdr.appendChild(createElementWithProps("w:left",attrs=bdrAttrs))
    divBdr.appendChild(createElementWithProps("w:bottom",attrs=bdrAttrs))
    divBdr.appendChild(createElementWithProps("w:right",attrs=bdrAttrs))
    div.appendChild(divBdr)
    divs.appendChild(div)
    webxml.appendChild(divs)
    webxml.appendChild(wxml.createElement("w:optimizeForBrowser"))
    webxml.appendChild(wxml.createElement("w:allowPNG"))
    wxml.appendChild(webxml)

    return wxml

def settings():
    sxml = Document()
    sett_attrs = {
        "mc:Ignorable": "w14 w15",
        "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "xmlns:o": "urn:schemas-microsoft-com:office:office",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
        "xmlns:v": "urn:schemas-microsoft-com:vml",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w10": "urn:schemas-microsoft-com:office:word",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml"
    }
    settxml = createElementWithProps("w:settings", attrs=sett_attrs)

    settxml.appendChild(createElementWithProps("w:zoom", attrs={"w:percent": "100"}))
    settxml.appendChild(createElementWithProps("w:proofState", attrs={"w:grammar": "clean", "w:spelling": "clean"}))
    settxml.appendChild(createElementWithProps("w:defaultTabStop", attrs={"w:val": "720"}))
    settxml.appendChild(createElementWithProps("w:characterSpacingControl", attrs={"w:val": "doNotCompress"}))

    compat = sxml.createElement("w:compat")
    compatSettings = (
        ("compatibilityMode", "http://schemas.microsoft.com/office/word", "15"),
        ("overrideTableStyleFontSizeAndJustification", "http://schemas.microsoft.com/office/word", "1"),
        ("enableOpenTypeFeatures", "http://schemas.microsoft.com/office/word", "1"),
        ("doNotFlipMirrorIndents", "http://schemas.microsoft.com/office/word", "1"),
        ("differentiateMultirowTableHeaders", "http://schemas.microsoft.com/office/word", "1")
    )
    for name, uri, val in compatSettings:
        cs_attrs = {"w:name": name, "w:uri": uri, "w:val": val}
        compat.appendChild(createElementWithProps("w:compatSetting", attrs=cs_attrs))
    settxml.appendChild(compat)

    # TODO: Actually generate this (what does it do?)
    rsids = sxml.createElement("w:rsids")
    rsid_data = (
        ("w:rsidRoot", "000F58CA"),
        ("w:rsid", "000F58CA"),
        ("w:rsid", "00286E1B"),
        ("w:rsid", "00382026")
    )
    for tag, val in rsid_data:
        rsids.appendChild(createElementWithProps(tag, attrs={"w:val": val}))
    settxml.appendChild(rsids)

    mathPr = sxml.createElement("w:mathPr")
    mathpr_data = (
        ("m:mathFont", "Cambria Math"),
        ("m:brkBin", "before"),
        ("m:brkBinSub", "--"),
        ("m:smallFrac", "0"),
        ("m:dispDef", None),
        ("m:lMargin", "0"),
        ("m:rMargin", "0"),
        ("m:defJc", "centerGroup"),
        ("m:wrapIndent", "1440"),
        ("m:intLim", "subSup"),
        ("m:naryLim", "undOvr")
    )
    for tag, val in mathpr_data:
        if val is None:
            attrs = {}
        else:
            attrs = {"w:val": val}
        mathPr.appendChild(createElementWithProps(tag, attrs=attrs))
    settxml.appendChild(mathPr)

    settxml.appendChild(createElementWithProps("w:themeFontLang", attrs={"w:val":"en-US"}))
    clrsm_attrs = {
        "w:accent1": "accent1",
        "w:accent2": "accent2",
        "w:accent3": "accent3",
        "w:accent4": "accent4",
        "w:accent5": "accent5",
        "w:accent6": "accent6",
        "w:bg1": "light1",
        "w:bg2": "light2",
        "w:followedHyperlink": "followedHyperlink",
        "w:hyperlink": "hyperlink",
        "w:t1": "dark1",
        "w:t2": "dark2"
    }
    settxml.appendChild(createElementWithProps("w:clrSchemeMapping", attrs=clrsm_attrs))

    shapeDefaults = sxml.createElement("w:shapeDefaults")
    shapeDefaults.appendChild(createElementWithProps("o:shapedefaults", attrs={"spidmax": "1026", "v:ext": "edit"}))
    shapelayout = createElementWithProps("o:shapelayout", attrs={"v:ext": "edit"})
    shapelayout.appendChild(createElementWithProps("o:idmap", attrs={"data": "1", "v:ext": "edit"}))
    shapeDefaults.appendChild(shapelayout)
    settxml.appendChild(shapeDefaults)

    settxml.appendChild(createElementWithProps("w:decimalSymbol", attrs={"w:val": "."}))
    settxml.appendChild(createElementWithProps("w:listSeparator", attrs={"w:val": ","}))
    settxml.appendChild(createElementWithProps("w15:chartTrackingRefBased"))
    # TODO: Actually generate this
    settxml.appendChild(createElementWithProps("w15:docId", attrs={"w15:val": "{49DA4288-3311-4E40-A92C-EB09F154294C}"}))

    sxml.appendChild(settxml)

    return sxml

def styles():
    pass

def fontTable():
    fxml = Document()
    font_attrs = {
        "mc:Ignorable": "w14 w15",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml"
    }
    fontxml = createElementWithProps("w:fonts", attrs=font_attrs)
    # TODO: Figure out where this font info comes from for further use
    font_data = {
        "Calibri": ("020F0502020204030204", "00", "swiss", "variable", "0000019F", "00000000", "E00002FF", "4000ACFF", "00000001", "00000000"),
        "Times New Roman": ("02020603050405020304", "00", "roman", "variable", "000001FF", "00000000", "E0002AFF", "C0007843", "00000009", "00000000"),
        "Calibri Light": ("020F0302020204030204", "00", "swiss", "variable", "0000019F", "00000000", "A00002EF", "4000207B", "00000000", "00000000")
    }
    for name, data in font_data.items():
        font = createElementWithProps("w:font", attrs={"w:name": name})
        font.appendChild(createElementWithProps("w:panose1", attrs={"w:val": data[0]}))
        font.appendChild(createElementWithProps("w:charset", attrs={"w:val": data[1]}))
        font.appendChild(createElementWithProps("w:family", attrs={"w:val": data[2]}))
        font.appendChild(createElementWithProps("w:pitch", attrs={"w:val": data[3]}))
        sigdata = {
            "w:csb0": data[4],
            "w:csb1": data[5],
            "w:usb0": data[6],
            "w:usb1": data[7],
            "w:usb2": data[8],
            "w:usb3": data[9]
        }
        font.appendChild(createElementWithProps("w:sig", attrs=sigdata))
        fontxml.appendChild(font)
    fxml.appendChild(fontxml)

    return fxml

def theme():
    pass

def makeAuxFiles(docx):
    aux = {}

    aux["[Content_Types].xml"] = contentTypes()
    aux["_rels/.rels"], aux["word/_rels/document.xml.rels"] = relationshipFiles()
    aux["docProps/core.xml"] = coreXML(docx)
    aux["docProps/app.xml"] = appXML()

    #aux["word/theme/theme1.xml"] = theme()
    aux["word/fontTable.xml"] = fontTable()
    aux["word/settings.xml"] = settings()
    #aux["word/styles.xml"] = styles()
    aux["word/webSettings.xml"] = webSettings()

    return aux