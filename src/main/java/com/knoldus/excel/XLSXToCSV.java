package com.knoldus.excel;

import java.io.InputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

class XLSXToCSV {

    private static ArrayList<String> contents = new ArrayList();
    private static ArrayList<ArrayList<String>> listOfRows = new ArrayList();
    private static PrintWriter outFile;
    private static String separator;
    private static int rowCount;
    private static int headerLength;
    private static boolean malformed;

    public ArrayList<ArrayList<String>> getListOfRows() {
        return listOfRows;
    }

    private void cleanUp() {
        this.malformed = false;
        this.rowCount = 0;
        this.headerLength = 0;
        this.contents.clear();
        this.getListOfRows().clear();
    }

    public void processOneSheet(String filename, String outPutFile, String separator) throws Exception {
        this.cleanUp();
        this.separator = separator;
        this.outFile = new PrintWriter(outPutFile);
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        InputStream sheet2 = r.getSheet("rId2");
        InputSource sheetSource = new InputSource(sheet2);
        parser.parse(sheetSource);
        sheet2.close();
        if (malformed) {
            outFile.close();
            throw new Exception("Malformed data.. check for missing cells");
        }
        outFile.close();
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser =
                XMLReaderFactory.createXMLReader(
                        "org.apache.xerces.parsers.SAXParser"
                );
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    private static class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            if (name.equals("c")) {
                int columnIndex = getColumnIndex(attributes.getValue("r"));
                if (columnIndex > 0 && contents.size() < columnIndex && columnIndex < headerLength -1) {
                    while (contents.size() != columnIndex) {
                        contents.add("");
                    }
                }
                String cellType = attributes.getValue("t");
                if (cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }
            lastContents = "";
        }

        public void endElement(String uri, String localName, String name)
                throws SAXException {
            if (nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                nextIsString = false;
            }

            if (name.equals("v") && !lastContents.equals("")) {
                contents.add(lastContents);
                if (rowCount == 0) {
                    headerLength++;
                }
            }

            if (name.equals("c") && lastContents.equals("") && contents.size() < headerLength) {
                contents.add("");
            }

            if (name.equals("row")) {
                for (int i = 0; i < contents.toArray().length; i++) {
                    if (contents.get(i).contains(",")) {
                        outFile.print("\"" + contents.get(i) + "\"");
                    } else {
                        outFile.print(contents.get(i));
                    }
                    if (i < contents.toArray().length - 1)
                        outFile.print(separator);
                }

                if (contents.size() < headerLength) {
                    int currentCellIndex = contents.size();
                    while (currentCellIndex < headerLength) {
                        contents.add("");
                        outFile.print(separator);
                        outFile.print("");
                        currentCellIndex++;
                    }
                }

                if (rowCount > 0) {
                    if (headerLength != contents.size()) {
                        malformed = true;
                    }
                }
                rowCount++;
                listOfRows.add(contents);
                outFile.append('\n');
                contents.clear();
            }
        }

        public void characters(char[] ch, int start, int length)
                throws SAXException {
            lastContents += new String(ch, start, length);
        }

        public int getColumnIndex(String cellReference) {
            String cellAlpha = "";
            for (int i = 0; i < cellReference.length(); i++) {
                if (Character.isDigit(cellReference.charAt(i)))
                    break;
                else
                    cellAlpha += cellReference.charAt(i);
            }
            return Alphabet.getNum(cellAlpha);
        }

        public enum Alphabet {
            A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z;
            public static int getNum(String targ) {
                return valueOf(targ).ordinal();
            }
        }
    }
}
