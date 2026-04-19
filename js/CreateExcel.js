const TEMPLATE_SHEET_NAME = "OS1";

async function createExcel() {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", createContentTypes());
    zip.folder("_rels").file(".rels", createRelsRels());
    zip.folder("xl").file("workbook.xml", createXlWorkbook());
    zip.folder("xl").folder("_rels").file("workbook.xml.rels", createXlRelsWorkbook());
    zip.folder("xl").folder("worksheets").file("sheet1.xml", createXlWorkSheetsSheet1());
    zip.folder("xl").file("styles.xml", createXlStyles());
    zip.folder("docProps").file("core.xml", createDocPropsCore());
    zip.folder("docProps").file("app.xml", createDocPropsApp());
    zip.folder("xl").folder("theme").file("theme1.xml", createXlThemeTheme1());
    zip.folder("xl").file("sharedStrings.xml", createXlSharedstrings());
    return await zip.generateAsync({ type: "blob" });
}

function createContentTypes() {
    const contentTypes = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
            <Default Extension="xml" ContentType="application/xml" />
            <Override PartName="/xl/workbook.xml"
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
            <Override PartName="/xl/worksheets/sheet1.xml"
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
            <Override PartName="/xl/theme/theme1.xml"
                ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />
            <Override PartName="/xl/styles.xml"
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
            <Override PartName="/xl/sharedStrings.xml"
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
            <Override PartName="/docProps/core.xml"
                ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
            <Override PartName="/docProps/app.xml"
                ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
        </Types>
    `.trim();
    return contentTypes;
}

function createRelsRels() {
    const rels = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                Target="docProps/app.xml" />
            <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                Target="docProps/core.xml" />
            <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="xl/workbook.xml" />
        </Relationships>
    `.trim();
    return rels;
}

function createXlWorkbook() {
    const workbook = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x15 xr xr6 xr10 xr2"
            xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
            xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
            xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
            xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
            <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="10603" />
            <workbookPr />
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                <mc:Choice Requires="x15">
                    <x15ac:absPath url="/Users/fndm90/Downloads/"
                        xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" />
                </mc:Choice>
            </mc:AlternateContent>
            <xr:revisionPtr revIDLastSave="0" documentId="13_ncr:1_{2C48E633-4116-E845-9E9B-181C66823B32}"
                xr6:coauthVersionLast="47" xr6:coauthVersionMax="47"
                xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}" />
            <bookViews>
                <workbookView xWindow="14300" yWindow="500" windowWidth="26660" windowHeight="25100"
                    xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}" />
            </bookViews>
            <sheets>
                <sheet name="${TEMPLATE_SHEET_NAME}" sheetId="1" r:id="rId1" />
            </sheets>
            <calcPr calcId="191029" />
        </workbook>
    `.trim();
    return workbook;
}

function createXlRelsWorkbook() {
    const workbookRels = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                Target="worksheets/sheet1.xml" />
            <Relationship Id="rId11"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
                Target="sharedStrings.xml" />
            <Relationship Id="rId10"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml" />
            <Relationship Id="rId9"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                Target="theme/theme1.xml" />
        </Relationships>
    `.trim();
    return workbookRels;
}

function createXlWorkSheetsSheet1() {
    const sheet1 = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac xr xr2 xr3"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
            xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
            xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
            xr:uid="{00000000-0001-0000-0000-000000000000}">
            <dimension ref="A1:H6" />
            <sheetViews>
                <sheetView tabSelected="1" zoomScaleNormal="100" workbookViewId="0">
                    <pane xSplit="1" ySplit="4" topLeftCell="B5" activePane="bottomRight" state="frozen" />
                    <selection activeCell="C5" sqref="C5" />
                    <selection pane="topRight" activeCell="C5" sqref="C5" />
                    <selection pane="bottomLeft" activeCell="C5" sqref="C5" />
                    <selection pane="bottomRight" />
                </sheetView>
            </sheetViews>
            <sheetFormatPr baseColWidth="10" defaultColWidth="11.1640625" defaultRowHeight="15"
                customHeight="1" />
            <cols>
                <col min="1" max="1" width="12.1640625" customWidth="1" />
                <col min="2" max="8" width="42.1640625" customWidth="1" />
            </cols>
            <sheetData>
                <row r="1" spans="1:8" ht="19.5" customHeight="1">
                    <c r="A1" s="1" t="s">
                        <v>0</v>
                    </c>
                    <c r="B1" s="2" t="s">
                        <v>1</v>
                    </c>
                    <c r="C1" s="3" />
                    <c r="D1" s="4" />
                    <c r="E1" s="2" t="s">
                        <v>2</v>
                    </c>
                    <c r="F1" s="10" />
                    <c r="G1" s="5" />
                    <c r="H1" s="17" />
                </row>
                <row r="2" spans="1:8" ht="28" thickBot="1">
                    <c r="A2" s="6" />
                    <c r="B2" s="7" t="s">
                        <v>3</v>
                    </c>
                    <c r="C2" s="14" />
                    <c r="D2" s="15" />
                    <c r="E2" s="15" />
                    <c r="F2" s="16" />
                    <c r="G2" s="5" />
                    <c r="H2" s="18" />
                </row>
                <row r="3" spans="1:8" ht="19.5" customHeight="1" thickBot="1">
                    <c r="A3" s="8" t="s">
                        <v>4</v>
                    </c>
                    <c r="B3" s="9" />
                    <c r="C3" s="9" />
                    <c r="D3" s="9" />
                    <c r="E3" s="9" />
                    <c r="F3" s="9" />
                    <c r="G3" s="9" />
                    <c r="H3" s="9" />
                </row>
                <row r="4" spans="1:8" ht="6" customHeight="1" thickTop="1">
                    <c r="A4" s="11" />
                    <c r="B4" s="11" />
                    <c r="C4" s="11" />
                    <c r="D4" s="11" />
                    <c r="E4" s="11" />
                    <c r="F4" s="11" />
                    <c r="G4" s="11" />
                    <c r="H4" s="11" />
                </row>
                <row r="5" spans="1:8" ht="24.75" customHeight="1">
                    <c r="A5" s="12" />
                    <c r="B5" s="12" />
                    <c r="C5" s="12" />
                    <c r="D5" s="12" />
                    <c r="E5" s="12" />
                    <c r="F5" s="12" />
                    <c r="G5" s="12" />
                    <c r="H5" s="12" />
                </row>
                <row r="6" spans="1:8" ht="409" customHeight="1">
                    <c r="A6" s="13" />
                    <c r="B6" s="12" />
                    <c r="C6" s="12" />
                    <c r="D6" s="12" />
                    <c r="E6" s="12" />
                    <c r="F6" s="12" />
                    <c r="G6" s="12" />
                    <c r="H6" s="12" />
                </row>
            </sheetData>
            <mergeCells count="1">
                <mergeCell ref="C2:F2" />
            </mergeCells>
            <phoneticPr fontId="5" />
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0" footer="0" />
            <pageSetup paperSize="9" orientation="portrait" />
        </worksheet>
    `.trim();
    return sheet1;
}

function createXlStyles() {
    const styles = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="x14ac x16r2 xr"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
            xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
            <fonts count="6">
                <font>
                    <sz val="12" />
                    <color theme="1" />
                    <name val="Calibri" />
                    <scheme val="minor" />
                </font>
                <font>
                    <b />
                    <sz val="13" />
                    <color theme="1" />
                    <name val="游ゴシック" />
                    <family val="3" />
                    <charset val="128" />
                </font>
                <font>
                    <b />
                    <sz val="16" />
                    <color theme="1" />
                    <name val="游ゴシック" />
                    <family val="3" />
                    <charset val="128" />
                </font>
                <font>
                    <sz val="14" />
                    <color theme="1" />
                    <name val="游ゴシック" />
                    <family val="3" />
                    <charset val="128" />
                </font>
                <font>
                    <sz val="12" />
                    <color theme="1" />
                    <name val="游ゴシック" />
                    <family val="3" />
                    <charset val="128" />
                </font>
                <font>
                    <sz val="6" />
                    <name val="Calibri" />
                    <family val="3" />
                    <charset val="128" />
                    <scheme val="minor" />
                </font>
            </fonts>
            <fills count="6">
                <fill>
                    <patternFill patternType="none" />
                </fill>
                <fill>
                    <patternFill patternType="gray125" />
                </fill>
                <fill>
                    <patternFill patternType="solid">
                        <fgColor rgb="FFBF9000" />
                        <bgColor rgb="FFBF9000" />
                    </patternFill>
                </fill>
                <fill>
                    <patternFill patternType="solid">
                        <fgColor theme="0" />
                        <bgColor theme="0" />
                    </patternFill>
                </fill>
                <fill>
                    <patternFill patternType="solid">
                        <fgColor rgb="FFFFD965" />
                        <bgColor rgb="FFFFD965" />
                    </patternFill>
                </fill>
                <fill>
                    <patternFill patternType="solid">
                        <fgColor rgb="FFFEF2CB" />
                        <bgColor rgb="FFFEF2CB" />
                    </patternFill>
                </fill>
            </fills>
            <borders count="14">
                <border>
                    <left />
                    <right />
                    <top />
                    <bottom />
                    <diagonal />
                </border>
                <border>
                    <left style="medium">
                        <color rgb="FF000000" />
                    </left>
                    <right style="medium">
                        <color rgb="FF000000" />
                    </right>
                    <top style="medium">
                        <color rgb="FF000000" />
                    </top>
                    <bottom style="thin">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left style="thin">
                        <color rgb="FF000000" />
                    </left>
                    <right style="thin">
                        <color rgb="FF000000" />
                    </right>
                    <top style="thin">
                        <color rgb="FF000000" />
                    </top>
                    <bottom style="thin">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left />
                    <right />
                    <top />
                    <bottom />
                    <diagonal />
                </border>
                <border>
                    <left style="medium">
                        <color rgb="FF000000" />
                    </left>
                    <right style="medium">
                        <color rgb="FF000000" />
                    </right>
                    <top />
                    <bottom style="medium">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left />
                    <right />
                    <top />
                    <bottom style="double">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left style="thin">
                        <color rgb="FF000000" />
                    </left>
                    <right style="thin">
                        <color rgb="FF000000" />
                    </right>
                    <top style="thin">
                        <color rgb="FF000000" />
                    </top>
                    <bottom style="double">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left style="thin">
                        <color rgb="FF000000" />
                    </left>
                    <right />
                    <top style="thin">
                        <color rgb="FF000000" />
                    </top>
                    <bottom style="thin">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left />
                    <right />
                    <top style="thin">
                        <color rgb="FF000000" />
                    </top>
                    <bottom style="thin">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left />
                    <right style="thin">
                        <color rgb="FF000000" />
                    </right>
                    <top style="thin">
                        <color rgb="FF000000" />
                    </top>
                    <bottom style="thin">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left style="thin">
                        <color indexed="64" />
                    </left>
                    <right style="thin">
                        <color indexed="64" />
                    </right>
                    <top style="thin">
                        <color indexed="64" />
                    </top>
                    <bottom style="thin">
                        <color indexed="64" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left />
                    <right />
                    <top style="double">
                        <color rgb="FF000000" />
                    </top>
                    <bottom />
                    <diagonal />
                </border>
                <border>
                    <left />
                    <right style="thin">
                        <color indexed="64" />
                    </right>
                    <top />
                    <bottom style="thin">
                        <color rgb="FF000000" />
                    </bottom>
                    <diagonal />
                </border>
                <border>
                    <left />
                    <right style="thin">
                        <color indexed="64" />
                    </right>
                    <top />
                    <bottom />
                    <diagonal />
                </border>
            </borders>
            <cellStyleXfs count="1">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
            </cellStyleXfs>
            <cellXfs count="19">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="center" vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="2" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="center" vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="3" fillId="3" borderId="3" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="3" borderId="3" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="2" borderId="3" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="3" borderId="4" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="2" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="center" vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="5" borderId="5" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="center" vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="5" borderId="6" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="center" vertical="center" />
                </xf>
                <xf numFmtId="14" fontId="3" fillId="3" borderId="2" xfId="0" applyNumberFormat="1"
                    applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="0" borderId="11" xfId="0" applyFont="1" applyBorder="1"
                    applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="0" borderId="10" xfId="0" applyFont="1" applyBorder="1"
                    applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="0" borderId="10" xfId="0" quotePrefix="1" applyFont="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="3" fillId="3" borderId="7" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="left" vertical="top" wrapText="1" />
                </xf>
                <xf numFmtId="0" fontId="3" fillId="3" borderId="8" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="left" vertical="top" wrapText="1" />
                </xf>
                <xf numFmtId="0" fontId="3" fillId="3" borderId="9" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment horizontal="left" vertical="top" wrapText="1" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="2" borderId="13" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
                <xf numFmtId="0" fontId="4" fillId="2" borderId="12" xfId="0" applyFont="1" applyFill="1"
                    applyBorder="1" applyAlignment="1">
                    <alignment vertical="center" />
                </xf>
            </cellXfs>
            <cellStyles count="1">
                <cellStyle name="標準" xfId="0" builtinId="0" />
            </cellStyles>
            <dxfs count="0" />
            <tableStyles count="0" defaultTableStyle="TableStyleMedium2"
                defaultPivotStyle="PivotStyleLight16" />
            <extLst>
                <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
                    xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                    <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1" />
                </ext>
                <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}"
                    xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                    <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1" />
                </ext>
            </extLst>
        </styleSheet>
    `.trim();
    return styles;
}

function createDocPropsCore() {
    const core = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties
            xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
            xmlns:dcmitype="http://purl.org/dc/dcmitype/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <dc:title></dc:title>
            <dc:subject></dc:subject>
            <dc:creator></dc:creator>
            <cp:keywords></cp:keywords>
            <dc:description></dc:description>
            <cp:lastModifiedBy>荒木 克也</cp:lastModifiedBy>
            <dcterms:created xsi:type="dcterms:W3CDTF">2020-09-01T01:00:58Z</dcterms:created>
            <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-22T01:52:33Z</dcterms:modified>
            <cp:category></cp:category>
        </cp:coreProperties>
    `.trim();
    return core;
}

function createDocPropsApp() {
    const app = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            <Application>Microsoft Macintosh Excel</Application>
            <DocSecurity>0</DocSecurity>
            <ScaleCrop>false</ScaleCrop>
            <HeadingPairs>
                <vt:vector size="2" baseType="variant">
                    <vt:variant>
                        <vt:lpstr>ワークシート</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>1</vt:i4>
                    </vt:variant>
                </vt:vector>
            </HeadingPairs>
            <TitlesOfParts>
                <vt:vector size="1" baseType="lpstr">
                    <vt:lpstr>${TEMPLATE_SHEET_NAME}</vt:lpstr>
                </vt:vector>
            </TitlesOfParts>
            <Manager></Manager>
            <Company></Company>
            <LinksUpToDate>false</LinksUpToDate>
            <SharedDoc>false</SharedDoc>
            <HyperlinkBase></HyperlinkBase>
            <HyperlinksChanged>false</HyperlinksChanged>
            <AppVersion>16.0300</AppVersion>
        </Properties>
    `.trim();
    return app;
}

function createXlThemeTheme1() {
    const theme1 = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Sheets">
            <a:themeElements>
                <a:clrScheme name="Sheets">
                    <a:dk1>
                        <a:srgbClr val="000000" />
                    </a:dk1>
                    <a:lt1>
                        <a:srgbClr val="FFFFFF" />
                    </a:lt1>
                    <a:dk2>
                        <a:srgbClr val="000000" />
                    </a:dk2>
                    <a:lt2>
                        <a:srgbClr val="FFFFFF" />
                    </a:lt2>
                    <a:accent1>
                        <a:srgbClr val="4472C4" />
                    </a:accent1>
                    <a:accent2>
                        <a:srgbClr val="ED7D31" />
                    </a:accent2>
                    <a:accent3>
                        <a:srgbClr val="A5A5A5" />
                    </a:accent3>
                    <a:accent4>
                        <a:srgbClr val="FFC000" />
                    </a:accent4>
                    <a:accent5>
                        <a:srgbClr val="5B9BD5" />
                    </a:accent5>
                    <a:accent6>
                        <a:srgbClr val="70AD47" />
                    </a:accent6>
                    <a:hlink>
                        <a:srgbClr val="0563C1" />
                    </a:hlink>
                    <a:folHlink>
                        <a:srgbClr val="0563C1" />
                    </a:folHlink>
                </a:clrScheme>
                <a:fontScheme name="Sheets">
                    <a:majorFont>
                        <a:latin typeface="Calibri" />
                        <a:ea typeface="Calibri" />
                        <a:cs typeface="Calibri" />
                    </a:majorFont>
                    <a:minorFont>
                        <a:latin typeface="Calibri" />
                        <a:ea typeface="Calibri" />
                        <a:cs typeface="Calibri" />
                    </a:minorFont>
                </a:fontScheme>
                <a:fmtScheme name="Office">
                    <a:fillStyleLst>
                        <a:solidFill>
                            <a:schemeClr val="phClr" />
                        </a:solidFill>
                        <a:gradFill rotWithShape="1">
                            <a:gsLst>
                                <a:gs pos="0">
                                    <a:schemeClr val="phClr">
                                        <a:lumMod val="110000" />
                                        <a:satMod val="105000" />
                                        <a:tint val="67000" />
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="50000">
                                    <a:schemeClr val="phClr">
                                        <a:lumMod val="105000" />
                                        <a:satMod val="103000" />
                                        <a:tint val="73000" />
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="100000">
                                    <a:schemeClr val="phClr">
                                        <a:lumMod val="105000" />
                                        <a:satMod val="109000" />
                                        <a:tint val="81000" />
                                    </a:schemeClr>
                                </a:gs>
                            </a:gsLst>
                            <a:lin ang="5400000" scaled="0" />
                        </a:gradFill>
                        <a:gradFill rotWithShape="1">
                            <a:gsLst>
                                <a:gs pos="0">
                                    <a:schemeClr val="phClr">
                                        <a:satMod val="103000" />
                                        <a:lumMod val="102000" />
                                        <a:tint val="94000" />
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="50000">
                                    <a:schemeClr val="phClr">
                                        <a:satMod val="110000" />
                                        <a:lumMod val="100000" />
                                        <a:shade val="100000" />
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="100000">
                                    <a:schemeClr val="phClr">
                                        <a:lumMod val="99000" />
                                        <a:satMod val="120000" />
                                        <a:shade val="78000" />
                                    </a:schemeClr>
                                </a:gs>
                            </a:gsLst>
                            <a:lin ang="5400000" scaled="0" />
                        </a:gradFill>
                    </a:fillStyleLst>
                    <a:lnStyleLst>
                        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                            <a:solidFill>
                                <a:schemeClr val="phClr" />
                            </a:solidFill>
                            <a:prstDash val="solid" />
                            <a:miter lim="800000" />
                        </a:ln>
                        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                            <a:solidFill>
                                <a:schemeClr val="phClr" />
                            </a:solidFill>
                            <a:prstDash val="solid" />
                            <a:miter lim="800000" />
                        </a:ln>
                        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                            <a:solidFill>
                                <a:schemeClr val="phClr" />
                            </a:solidFill>
                            <a:prstDash val="solid" />
                            <a:miter lim="800000" />
                        </a:ln>
                    </a:lnStyleLst>
                    <a:effectStyleLst>
                        <a:effectStyle>
                            <a:effectLst />
                        </a:effectStyle>
                        <a:effectStyle>
                            <a:effectLst />
                        </a:effectStyle>
                        <a:effectStyle>
                            <a:effectLst>
                                <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr"
                                    rotWithShape="0">
                                    <a:srgbClr val="000000">
                                        <a:alpha val="63000" />
                                    </a:srgbClr>
                                </a:outerShdw>
                            </a:effectLst>
                        </a:effectStyle>
                    </a:effectStyleLst>
                    <a:bgFillStyleLst>
                        <a:solidFill>
                            <a:schemeClr val="phClr" />
                        </a:solidFill>
                        <a:solidFill>
                            <a:schemeClr val="phClr">
                                <a:tint val="95000" />
                                <a:satMod val="170000" />
                            </a:schemeClr>
                        </a:solidFill>
                        <a:gradFill rotWithShape="1">
                            <a:gsLst>
                                <a:gs pos="0">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="93000" />
                                        <a:satMod val="150000" />
                                        <a:shade val="98000" />
                                        <a:lumMod val="102000" />
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="50000">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="98000" />
                                        <a:satMod val="130000" />
                                        <a:shade val="90000" />
                                        <a:lumMod val="103000" />
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="100000">
                                    <a:schemeClr val="phClr">
                                        <a:shade val="63000" />
                                        <a:satMod val="120000" />
                                    </a:schemeClr>
                                </a:gs>
                            </a:gsLst>
                            <a:lin ang="5400000" scaled="0" />
                        </a:gradFill>
                    </a:bgFillStyleLst>
                </a:fmtScheme>
            </a:themeElements>
            <a:objectDefaults />
            <a:extraClrSchemeLst />
        </a:theme>
    `.trim();
    return theme1;
}

function createXlSharedstrings() {
    const sharedStrings = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
            <si>
                <t>案件/バージョン</t>
            </si>
            <si>
                <t>試験仕様書名</t>
            </si>
            <si>
                <t>実施日時</t>
            </si>
            <si>
                <t>その他</t>
            </si>
            <si>
                <t>項目↓/手順→</t>
            </si>
        </sst>
    `.trim();
    return sharedStrings
}
