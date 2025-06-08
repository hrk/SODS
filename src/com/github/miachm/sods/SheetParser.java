package com.github.miachm.sods;

import java.time.LocalDateTime;
import java.time.format.DateTimeParseException;
import java.util.*;

class SheetParser {
    private static final int BUGGED_COUNT = 10 * 1000;
    private final Sheet sheet;
    private final StylesParser stylesParser;
    private final Map<Integer, Style> columnDefaultStyles = new HashMap<>();
    private final Map<Integer, Style> rowDefaultStyles = new HashMap<>();
    private final Set<Pair<Vector, Vector>> groupCells = new HashSet<>();

    public SheetParser(Sheet sheet, StylesParser stylesParser) {
        this.sheet = sheet;
        this.stylesParser = stylesParser;
    }

    public void parseSheet(XmlReaderInstance reader) {
        String tableStyleName = reader.getAttribValue("table:style-name");
        if (tableStyleName != null) setTableStyles(tableStyleName);

        String protectedSheet = reader.getAttribValue("table:protected");
        if (protectedSheet != null) {
            String algorithm = reader.getAttribValue("table:protection-key-digest-algorithm");
            if (algorithm == null) algorithm = "http://www.w3.org/2000/09/xmldsig#sha1";
            String protectedKey = reader.getAttribValue("table:protection-key");
            sheet.setRawPassword(protectedKey, algorithm);
        }

        int rowCount = 0;
        groupCells.clear();

        while (reader.hasNext()) {
            XmlReaderInstance instance = reader.nextElement("table:table-column", "table:table-row");
            if (instance == null) break;

            String styleName = instance.getAttribValue("table:default-cell-style-name");
            Style style = styleName != null ? stylesParser.getCellStyle(styleName) : null;

            if (instance.getTag().equals("table:table-column")) {
                parseColumnProperties(instance, style);
            } else if (instance.getTag().equals("table:table-row")) {
                if (style != null) rowDefaultStyles.put(rowCount, style);

                int numRows = 1;
                String numRowsStr = instance.getAttribValue("table:number-rows-repeated");
                if (numRowsStr != null) {
                    try {
                        numRows = Integer.parseInt(numRowsStr);
                        if (numRows > BUGGED_COUNT) continue;
                    } catch (NumberFormatException ignored) {}
                }

                sheet.appendRows(numRows);

                String visibility = instance.getAttribValue("table:visibility");
                if ("collapse".equals(visibility)) sheet.hideRows(sheet.getMaxRows() - numRows, numRows);

                String rowStyleName = instance.getAttribValue("table:style-name");
                if (rowStyleName != null) {
                    RowStyle rowStyle = stylesParser.getRowStyle(rowStyleName);
                    if (rowStyle != null) sheet.setRowHeights(sheet.getMaxRows() - numRows, numRows, rowStyle.getHeight());
                }

                processCells(instance, numRows);
                rowCount += numRows;
            }
        }

        for (Pair<Vector, Vector> pair : groupCells) {
            Vector cord = pair.first;
            Vector length = pair.second;
            Range range = sheet.getRange(cord.getX(), cord.getY(), length.getX(), length.getY());
            range.merge();
        }
    }

    private void setTableStyles(String tableStyleName) {
        TableStyle style = stylesParser.getTableStyle(tableStyleName);
        if (style != null && style.isHidden()) sheet.hideSheet();
    }

    private void parseColumnProperties(XmlReaderInstance instance, Style style) {
        boolean areHidden = "collapse".equals(instance.getAttribValue("table:visibility"));
        int numColumns = 1;
        String columnsRepeated = instance.getAttribValue("table:number-columns-repeated");
        if (columnsRepeated != null) {
            numColumns = Integer.parseInt(columnsRepeated);
            if (numColumns > BUGGED_COUNT) return;
        }

        int index = sheet.getMaxColumns();
        sheet.appendColumns(numColumns);

        if (style != null && !style.isDefault()) {
            for (int j = index; j < index + numColumns; j++) {
                sheet.setDefaultColumnCellStyle(j, style);
                columnDefaultStyles.put(j, style);
            }
        }

        if (areHidden) sheet.hideColumns(index, numColumns);

        String columnStyleName = instance.getAttribValue("table:style-name");
        if (columnStyleName != null) {
            ColumnStyle columnStyle = stylesParser.getColumnStyle(columnStyleName);
            if (columnStyle != null) sheet.setColumnWidths(sheet.getMaxColumns() - numColumns, numColumns, columnStyle.getWidth());
        }
    }

    private void processCells(XmlReaderInstance reader, int numberRowsRepeated) {
        int column = 0;
        while (reader.hasNext()) {
            int numberColumnsRepeated = 1;
            Object lastCellValue = null;
            Style lastStyle = null;

            XmlReaderInstance instance = reader.nextElement("table:table-cell", "table:covered-table-cell");
            if (instance == null) break;

            if (instance.getTag().equals("table:covered-table-cell")) {
                String numColumnsRepeated = instance.getAttribValue("table:number-columns-repeated");
                column += numColumnsRepeated == null ? 1 : Integer.parseInt(numColumnsRepeated);
                continue;
            }

            int rows = 1, columns = 1;
            String rowsSpanned = instance.getAttribValue("table:number-rows-spanned");
            if (rowsSpanned != null) rows = Integer.parseInt(rowsSpanned);
            String columnsSpanned = instance.getAttribValue("table:number-columns-spanned");
            if (columnsSpanned != null) columns = Integer.parseInt(columnsSpanned);

            if (numberRowsRepeated == 1 && (rows != 1 || columns != 1)) {
                Pair<Vector, Vector> pair = new Pair<>();
                pair.first = new Vector(sheet.getMaxRows() - 1, column);
                pair.second = new Vector(rows, columns);
                groupCells.add(pair);
            }

            int positionX = sheet.getMaxRows() - numberRowsRepeated;
            int positionY = column;

            OfficeValueType valueType = OfficeValueType.ofReader(instance);
            Object value = valueType.read(instance);

            String raw = instance.getAttribValue("table:number-columns-repeated");
            if (raw != null) {
                numberColumnsRepeated = Integer.parseInt(raw);
                if (numberColumnsRepeated > BUGGED_COUNT) continue;
            }

            if (positionY + numberColumnsRepeated > sheet.getMaxColumns()) {
                sheet.appendColumns(positionY + numberColumnsRepeated - sheet.getMaxColumns());
            }

            Range range = sheet.getRange(positionX, positionY, numberRowsRepeated, numberColumnsRepeated);

            String formula = instance.getAttribValue("table:formula");
            if (formula != null) range.setFormula(formula);
            range.setValue(value);

            Style style = stylesParser.getCellStyle(instance.getAttribValue("table:style-name"));
            if (style == null) style = columnDefaultStyles.get(column);
            if (style == null) style = rowDefaultStyles.get(sheet.getMaxRows() - 1);
            if (style != null && !style.isDefault()) range.setStyle(style);

            readCellText(instance, range);
            column += numberColumnsRepeated;
        }
    }

    private void readCellText(XmlReaderInstance cellReader, Range range) {
        StringBuffer s = new StringBuffer();
        boolean firstTextElement = true;

        XmlReaderInstance textElement;
        while ((textElement = cellReader.nextElement("text:p", "text:h", "office:annotation")) != null) {
            if (textElement.getTag().equals("office:annotation")) {
                range.setAnnotation(getOfficeAnnotation(textElement));
                continue;
            }

            if (firstTextElement) firstTextElement = false;
            else s.append("\n");

            XmlReaderInstance spanElement;
            while ((spanElement = textElement.nextElement("text:s", XmlReaderInstance.CHARACTERS)) != null) {
                if (spanElement.getTag().equals("text:s")) {
                    int num = 1;
                    String attrib = spanElement.getAttribValue("text:c");
                    if (attrib != null && !attrib.isEmpty()) {
                        try {
                            num = Integer.parseInt(attrib);
                        } catch (NumberFormatException e) {
                            System.err.println("Invalid number of characters: " + attrib);
                        }
                    }
                    while (num-- > 0) s.append(" ");
                }

                String spanContent = spanElement.getContent();
                if (spanContent != null) s.append(spanContent);
            }
        }

        Object value = range.getValue();
        if (s.length() > 0 && (value == null || value instanceof String)) {
            range.setValue(s.toString());
        }
    }

    private OfficeAnnotation getOfficeAnnotation(XmlReaderInstance reader) {
        OfficeAnnotationBuilder annotation = new OfficeAnnotationBuilder();
        StringBuilder msg = new StringBuilder();

        while (reader.hasNext()) {
            XmlReaderInstance instance = reader.nextElement("dc:date", "text:p");
            if (instance == null) break;

            if (instance.getTag().equals("dc:date")) {
                instance = instance.nextElement(XmlReaderInstance.CHARACTERS);
                if (instance != null) {
                    String content = instance.getContent();
                    try {
                        if (content != null) annotation.setLastModified(LocalDateTime.parse(content));
                    } catch (DateTimeParseException e) {
                        System.err.println("DATE INVALID IN OFFICE ANNOTATION");
                    }
                }
            } else if (instance.getTag().equals("text:p")) {
                instance = instance.nextElement(XmlReaderInstance.CHARACTERS);
                if (msg.length() > 0) msg.append("\n");
                if (instance != null) msg.append(instance.getContent());
            }
        }

        annotation.setMsg(msg.toString());
        return annotation.build();
    }
}