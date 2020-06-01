package org.apache.poi.xssf.eventusermodel;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Optional;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 * A rudimentary XLSX -> CSV processor
 * based on XLS2CSVmra by Nick Burch from
 * package org.apache.poi.hssf.eventusermodel.examples.
 * This is an attempt to demonstrate the same thing using XSSF.
 * Unlike the HSSF version, this one completely ignores missing rows.
 */
public class XLSX2CSV {

    /**
     * The type of the data value is indicated by an attribute on
     * the cell element; the value is in a "v" element within the cell.
     */
    enum xssfDataType {
        BOOL,
        DATE,
        DATETIME,
        FORMULA,
        SSTINDEX,
        TIME,
        NUMBER,
    }

    /**
     * Derived from http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
     */
    class MyXSSFSheetHandler extends DefaultHandler {
        public int currentRow = -1;

        /** Table with unique strings */
        private ReadOnlySharedStringsTable sharedStringsTable;

        /** Destination for data */
        private final PrintStream output;

        /** Number of columns to read starting with leftmost */
        private final int minColumnCount;

        // Runtime
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("M/d/yyyy");
        SimpleDateFormat simpleTimeFormat = new SimpleDateFormat("hh:mm:ss a");

        // Set when V start element is seen
        private boolean vIsOpen;

        // Set when cell start element is seen;
        // used when cell close element is seen.
        private xssfDataType nextDataType;

        private int thisColumn = -1;
        // The last column printed to the output stream
        private int lastColumnNumber = -1;

        private StringBuffer contents;

        /**
         *
         * @param sst
         * @param cols
         * @param target
         */
        public MyXSSFSheetHandler(
                ReadOnlySharedStringsTable sst,
                int cols,
                PrintStream target) {
            this.sharedStringsTable = sst;
            this.minColumnCount = cols;
            this.output = target;
            this.contents = new StringBuffer();
            this.nextDataType = xssfDataType.NUMBER;
        }

        /*
         * (non-Javadoc)
         * @see org.xml.sax.helpers.DefaultHandler#startElement(java.lang.String, java.lang.String,
java.lang.String, org.xml.sax.Attributes)
         */
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            if("row".equals(name)) {
                // Get the cell reference
                String r = attributes.getValue("r");
                int firstDigit = -1;
                for (int c = 0; c < r.length(); ++c) {
                    if (Character.isDigit(r.charAt(c))) {
                        firstDigit = c;
                        break;
                    }
                }
                currentRow = Integer.parseInt(r.substring(firstDigit));
            }
//            System.out.println(firstRow.orElse(1) + "\t" + currentRow);
            if (firstRow.orElse(1) > currentRow) { return; }
            if (lastRow.isPresent() && lastRow.get() < currentRow) { return; }
            // c => cell
            if ("c".equals(name)) {
                // Get the cell reference
                String r = attributes.getValue("r");
                int firstDigit = -1;
                for (int c = 0; c < r.length(); ++c) {
                    if (Character.isDigit(r.charAt(c))) {
                        firstDigit = c;
                        break;
                    }
                }

                thisColumn = nameToColumn(r.substring(0, firstDigit));

                // Figure out if the value is an index in the SST
                // or something else.
                String cellType = attributes.getValue("t");
                String cellSomething = attributes.getValue("s");
                if ("b".equals(cellType))
                    nextDataType = xssfDataType.BOOL;
                else if ("e".equals(cellType))
                    nextDataType = xssfDataType.FORMULA;
                else if ("s".equals(cellType))
                    nextDataType = xssfDataType.SSTINDEX;
                else if ("2".equals(cellSomething))
                    nextDataType = xssfDataType.DATE;
                else if ("3".equals(cellSomething))
                    nextDataType = xssfDataType.TIME;
                else if ("4".equals(cellSomething))
                    nextDataType = xssfDataType.DATETIME;
                else
                    nextDataType = xssfDataType.NUMBER;
            }
            else if ("v".equals(name)) {
                vIsOpen = true;
                // Clear contents cache
                contents.setLength(0);
            }
        }

        /*
         * (non-Javadoc)
         * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String, java.lang.String,
java.lang.String)
         */
        public void endElement(String uri, String localName, String name)
                throws SAXException {

            String thisStr = null;
            if (firstRow.isPresent() && currentRow < firstRow.get()) { return; }
            if (lastRow.isPresent() && lastRow.get() < currentRow) { return; }
            // v => contents of a cell
            if ("v".equals(name)) {
                // Process the value contents as required.
                // Do now, as characters() may be called more than once
                switch(nextDataType) {

                    case BOOL:
                        char first = contents.charAt(0);
                        thisStr = first == '0' ? "FALSE" : "TRUE";
                        break;

                    case DATE:
                        // The value is actually an integer
                        long daysSince = Long.parseLong(contents.toString());
                        Date d = DateUtil.getJavaDate(daysSince);
                        thisStr = '"' + simpleDateFormat.format(d) + '"';
                        break;

                    case DATETIME:
                        // Days to left of decimal, seconds (?) to right of decimal.
//                        Date dt = DateUtil.getJavaDate(Double.parseDouble(contents.toString()));
                        thisStr = '"' + contents.toString() + '"';
                        break;

                    case SSTINDEX:
                        String sstIndex = contents.toString();
                        try {
                            int idx = Integer.parseInt(sstIndex);
                            XSSFRichTextString rts = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                            thisStr = '"' + rts.toString() + '"';
                        }
                        catch (NumberFormatException ex) {
                            output.println("Pgmr err, lastContents is not int: " + sstIndex);
                        }
                        break;

                    case TIME:
                        Date t = DateUtil.getJavaDate(Double.parseDouble(contents.toString()));
                        thisStr = '"' + simpleTimeFormat.format(t) + '"';
                        break;

                    case FORMULA:
                        // A formula could result in a string value,
                        // so always add doublequote characters.
                        thisStr = '"' + contents.toString() + '"';
                        break;

                    case NUMBER:
                        thisStr = contents.toString();
                        break;

                    default:
                        thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                        break;
                }

                // Output after we've seen the string contents
                // Emit commas for any fields that were missing on this row
                if(lastColumnNumber == -1) { lastColumnNumber = 0; }
                for (int i = lastColumnNumber; i < thisColumn; ++i)
                    output.print(',');

                // Might be the empty string.
                output.print(thisStr);

                // Update column
                if(thisColumn > -1)
                    lastColumnNumber = thisColumn;

            }
            else if("row".equals(name)) {

                // Print out any missing commas if needed
                if(minColumns > 0) {
                    // Columns are 0 based
                    if(lastColumnNumber == -1) { lastColumnNumber = 0; }
                    for(int i=lastColumnNumber; i<(this.minColumnCount); i++) {
                        output.print(',');
                    }
                }

                // We're onto a new row
                output.println();
                lastColumnNumber = -1;
            }

        }

        /**
         * Captures characters only if a v(alue) element is open.
         */
        public void characters(char[] ch, int start, int length)
                throws SAXException {
            if (vIsOpen)
                contents.append(ch, start, length);
        }

        /**
         * Converts an Excel column name like "C" to a zero-based index.
         * @param name
         * @return Index corresponding to the specified name
         */
        private int nameToColumn(String name) {
            int column = -1;
//            System.out.println(name);
            for (int i = 0; i < name.length(); ++i) {
                int c = name.charAt(i);
                column = (column + 1) * 26 + c - 'A';
            }
            return column;
        }

    }

    ///////////////////////////////////////

    private OPCPackage xlsxPackage;
    private int minColumns;
    private PrintStream output;
    private Optional<Integer> firstRow;
    private Optional<Integer> lastRow;
    private Optional<String> sheetRegExp;

    /**
     * Creates a new XLSX -> CSV converter
     * Javadoc says I should use OPCPackage instead of Package, but OPCPackage
     * was not available in the POI 3.5-beta5 build I had at the time.
     *
     * @param pkg The XLSX package to process
     * @param output The PrintStream to output the CSV to
     * @param minColumns The minimum number of columns to output, or -1 for no minimum
     */
    public XLSX2CSV(OPCPackage pkg, PrintStream output, int minColumns) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.minColumns = minColumns;
    }

    public XLSX2CSV(OPCPackage pkg, PrintStream output, int minColumns,
                    Optional<Integer> firstRow, Optional<Integer> lastRow, Optional<String> sheetRegExp) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.minColumns = minColumns;
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.sheetRegExp = sheetRegExp;
    }

    /**
     * @param sst
     * @param sheetInputStream
     */
    public void processSheet(ReadOnlySharedStringsTable sst, InputStream sheetInputStream)
            throws IOException, ParserConfigurationException, SAXException {

        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader sheetParser = saxParser.getXMLReader();
        ContentHandler handler = new MyXSSFSheetHandler(sst, this.minColumns, this.output);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
    }

    /**
     * Initiates the processing of the XLS file to CSV
     * @throws OpenXML4JException
     */
    public void process()
            throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {

        ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator)xssfReader.getSheetsData();
        int index = 0;
        while (iter.hasNext()) {
            InputStream stream = iter.next();
            String sheetName = iter.getSheetName();
            this.output.println();
            this.output.println(sheetName + " [index=" + index + "]:");
            processSheet(sst, stream);
            // stream.close();
            ++index;
        }
    }

    public static void main(String[] args) throws Exception {
        if(args.length < 1) {
            System.err.println("Use:");
            System.err.println("  XLSX2CSV <xlsx file> [min columns]");
            System.exit(1);
        }

        File xlsxFile = new File(args[0]);
        if (! xlsxFile.exists()) {
            System.err.println("Not found or not a file: " + xlsxFile.getPath());
            System.exit(1);
        }

        int minColumns = -1;
        if(args.length >= 2) {
            minColumns = Integer.parseInt(args[1]);
        }

        // If no log4j configuration is provided, these messages appear:
        //   log4j:WARN No appenders could be found for logger (org.openxml4j.opc).
        //   log4j:WARN Please initialize the log4j system properly.
        // If only the BasicConfigurator.configure() is done, these messages appear:
        //   0 [main] DEBUG org.openxml4j.opc  - Parsing relationship: /xl/_rels/workbook.xml.rels
        //  46 [main] DEBUG org.openxml4j.opc  - Parsing relationship: /_rels/.rels
        // Added the call to setLevel() to turn these off, now I see nothing.

        BasicConfigurator.configure();
        Logger.getRootLogger().setLevel(Level.INFO);

        // The package open is instantaneous, as it should be.
        OPCPackage p = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ);
        Optional<Integer> firstRow = Optional.ofNullable(5);
        Optional<Integer> lastRow = Optional.ofNullable(13);
        Optional<String> sheetRegExp = Optional.ofNullable("");
        XLSX2CSV xlsx2csv = new XLSX2CSV(p, new PrintStream(new BufferedOutputStream(new FileOutputStream("three_test.txt")), true), minColumns, firstRow, lastRow, sheetRegExp);
        xlsx2csv.process();
        // Want to call close() here, but the package is open for read,
        // so it's not necessary, and it complains if I do call it!
        p.revert();
    }

}