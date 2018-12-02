/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package excel.streaming;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.util.CellReference;

import java.io.*;
import java.util.Iterator;

/**
 * Initially copied from BigGridDemo "SpreadsheetWriter".
 * Unlike the original code which wrote the entire document,
 * this class only writes the "sheetData" document fragment
 * so that it was renamed to "SheetDataWriter"
 */
public class SheetDataWriter {
    private int _rownum;
    private boolean _rowContainedNullCells = false;
    int _numberOfFlushedRows;
    int _lowestIndexOfFlushedRows; // meaningful only of _numberOfFlushedRows>0
    int _numberOfCellsOfLastFlushedRow; // meaningful only of _numberOfFlushedRows>0

    public SheetDataWriter() {
    }

    public int getNumberOfFlushedRows() {
        return _numberOfFlushedRows;
    }

    public int getNumberOfCellsOfLastFlushedRow() {
        return _numberOfCellsOfLastFlushedRow;
    }

    public int getLowestIndexOfFlushedRows() {
        return _lowestIndexOfFlushedRows;
    }

    /**
     * Write a row to the file
     *
     * @param rownum 0-based row number
     * @param row    a row
     */
    public void writeRow(int rownum, StreamRow row, Writer out) throws IOException {
        if (_numberOfFlushedRows == 0)
            _lowestIndexOfFlushedRows = rownum;
        _numberOfCellsOfLastFlushedRow = row.getLastCellNum();
        _numberOfFlushedRows++;
        beginRow(rownum, row, out);
        Iterator<Cell> cells = row.allCellsIterator();
        int columnIndex = 0;
        while (cells.hasNext()) {
            writeCell(columnIndex++, cells.next(), out);
        }
        endRow(out);
    }

    void beginRow(int rownum, StreamRow row, Writer out) throws IOException {
        out.write("<row r=\"" + (rownum + 1) + "\"");
        if (row.hasCustomHeight())
            out.write(" customHeight=\"true\"  ht=\"" + row.getHeightInPoints() + "\"");
        if (row.getZeroHeight())
            out.write(" hidden=\"true\"");
        if (row.isFormatted()) {
            out.write(" s=\"" + row._style + "\"");
            out.write(" customFormat=\"1\"");
        }
        if (row.getOutlineLevel() != 0) {
            out.write(" outlineLevel=\"" + row.getOutlineLevel() + "\"");
        }
        out.write(">\n");
        this._rownum = rownum;
        _rowContainedNullCells = false;
    }

    void endRow(Writer out) throws IOException {
        out.write("</row>\n");
    }

    public void writeCell(int columnIndex, Cell cell, Writer out) throws IOException {
        if (cell == null) {
            _rowContainedNullCells = true;
            return;
        }
        String ref = new CellReference(_rownum, columnIndex).formatAsString();
        out.write("<c r=\"" + ref + "\"");
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle.getIndex() != 0) out.write(" s=\"" + cellStyle.getIndex() + "\"");
        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_BLANK: {
                out.write(">");
                break;
            }
            case Cell.CELL_TYPE_FORMULA: {
                out.write(">");
                out.write("<f>");
                outputQuotedString(cell.getCellFormula(), out);
                out.write("</f>");
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        double nval = cell.getNumericCellValue();
                        if (!Double.isNaN(nval)) {
                            out.write("<v>" + nval + "</v>");
                        }
                        break;
                }
                break;
            }
            case Cell.CELL_TYPE_STRING: {
                out.write(" t=\"inlineStr\">");
                out.write("<is><t");
                if(hasLeadingTrailingSpaces(cell.getStringCellValue())) {
                    out.write(" xml:space=\"preserve\"");
                }
                out.write(">");
                outputQuotedString(cell.getStringCellValue(), out);
                out.write("</t></is>");
                break;
            }
            case Cell.CELL_TYPE_NUMERIC: {
                out.write(" t=\"n\">");
                out.write("<v>" + cell.getNumericCellValue() + "</v>");
                break;
            }
            case Cell.CELL_TYPE_BOOLEAN: {
                out.write(" t=\"b\">");
                out.write("<v>" + (cell.getBooleanCellValue() ? "1" : "0") + "</v>");
                break;
            }
            case Cell.CELL_TYPE_ERROR: {
                FormulaError error = FormulaError.forInt(cell.getErrorCellValue());

                out.write(" t=\"e\">");
                out.write("<v>" + error.getString() + "</v>");
                break;
            }
            default: {
                throw new RuntimeException("Huh?");
            }
        }
        out.write("</c>");
    }


    /**
     * @return  whether the string has leading / trailing spaces that
     *  need to be preserved with the xml:space=\"preserve\" attribute
     */
    boolean hasLeadingTrailingSpaces(String str) {
        if (str != null && str.length() > 0) {
            char firstChar = str.charAt(0);
            char lastChar  = str.charAt(str.length() - 1);
            return Character.isWhitespace(firstChar) || Character.isWhitespace(lastChar) ;
        }
        return false;
    }

    //Taken from jdk1.3/src/javax/swing/text/html/HTMLWriter.java
    protected void outputQuotedString(String s, Writer _out) throws IOException {
        if (s == null || s.length() == 0) {
            return;
        }

        char[] chars = s.toCharArray();
        int last = 0;
        int length = s.length();
        for (int counter = 0; counter < length; counter++) {
            char c = chars[counter];
            switch (c) {
                case '<':
                    if (counter > last) {
                        _out.write(chars, last, counter - last);
                    }
                    last = counter + 1;
                    _out.write("&lt;");
                    break;
                case '>':
                    if (counter > last) {
                        _out.write(chars, last, counter - last);
                    }
                    last = counter + 1;
                    _out.write("&gt;");
                    break;
                case '&':
                    if (counter > last) {
                        _out.write(chars, last, counter - last);
                    }
                    last = counter + 1;
                    _out.write("&amp;");
                    break;
                case '"':
                    if (counter > last) {
                        _out.write(chars, last, counter - last);
                    }
                    last = counter + 1;
                    _out.write("&quot;");
                    break;
                // Special characters
                case '\n':
                case '\r':
                    if (counter > last) {
                        _out.write(chars, last, counter - last);
                    }
                    _out.write("&#xa;");
                    last = counter + 1;
                    break;
                case '\t':
                    if (counter > last) {
                        _out.write(chars, last, counter - last);
                    }
                    _out.write("&#x9;");
                    last = counter + 1;
                    break;
                case 0xa0:
                    if (counter > last) {
                        _out.write(chars, last, counter - last);
                    }
                    _out.write("&#xa0;");
                    last = counter + 1;
                    break;
                default:
                    // YK: XmlBeans silently replaces all ISO control characters ( < 32) with question marks.
                    // the same rule applies to unicode surrogates and "not a character" symbols.
                    if( c < ' ' || Character.isLowSurrogate(c) || Character.isHighSurrogate(c) ||
                            ('\uFFFE' <= c && c <= '\uFFFF')) {
                        if (counter > last) {
                            _out.write(chars, last, counter - last);
                        }
                        _out.write('?');
                        last = counter + 1;
                    }
                    else if (c > 127) {
                        if (counter > last) {
                            _out.write(chars, last, counter - last);
                        }
                        last = counter + 1;
                        // If the character is outside of ascii, write the
                        // numeric value.
                        _out.write("&#");
                        _out.write(String.valueOf((int) c));
                        _out.write(";");
                    }
                    break;
            }
        }
        if (last < length) {
            _out.write(chars, last, length - last);
        }
    }

}
