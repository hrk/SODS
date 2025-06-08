package com.github.miachm.sods;

class SpreadsheetParser {
    private final StylesParser stylesParser;
    private final SpreadSheet spread;

    public SpreadsheetParser(StylesParser stylesParser, SpreadSheet spread) {
        this.stylesParser = stylesParser;
        this.spread = spread;
    }

    public void parseContent(XmlReaderInstance bodyInstance) {
        if (bodyInstance == null) return;
        XmlReaderInstance spreadsheetInstance = bodyInstance.nextElement("office:spreadsheet");
        if (spreadsheetInstance != null) {
            while (spreadsheetInstance.hasNext()) {
                XmlReaderInstance tableInstance = spreadsheetInstance.nextElement("table:table");
                if (tableInstance != null) {
                    String name = tableInstance.getAttribValue("table:name");
                    Sheet sheet = new Sheet(name, 0, 0);
                    SheetParser sheetParser = new SheetParser(sheet, stylesParser);
                    sheetParser.parseSheet(tableInstance);
                    spread.appendSheet(sheet);
                }
            }
        }
    }
}