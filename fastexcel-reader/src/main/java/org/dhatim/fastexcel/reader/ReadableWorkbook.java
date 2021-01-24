/*
 * Copyright 2016 Dhatim.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.dhatim.fastexcel.reader;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class ReadableWorkbook implements Closeable {

    private final OPCPackage pkg;
    private final SST sst;
    private final static XMLInputFactory factory = defaultXmlInputFactory();

    private boolean date1904;
    private final List<Sheet> sheets = new ArrayList<>();

    public ReadableWorkbook(File inputFile) throws IOException {
        this(OPCPackage.open(inputFile, factory));
    }

    /**
     * Note: will load the whole xlsx file into memory,
     * (but will not uncompress it in memory)
     */
    public ReadableWorkbook(InputStream inputStream) throws IOException {
        this(OPCPackage.open(inputStream, factory));
    }

    private ReadableWorkbook(OPCPackage pkg) throws IOException {

        try {
            this.pkg = pkg;
            sst = SST.fromInputStream(factory, pkg.getSharedStrings());
        } catch (XMLStreamException e) {
            throw new ExcelReaderException(e);
        }

        try (SimpleXmlReader workbookReader = new SimpleXmlReader(factory, pkg.getWorkbookContent())) {
            readWorkbook(workbookReader);
        } catch (XMLStreamException e) {
            throw new ExcelReaderException(e);
        }
    }

    private static XMLInputFactory defaultXmlInputFactory() {
        XMLInputFactory factory = XMLInputFactory.newInstance();
        // To prevent XML External Entity (XXE) attacks
        factory.setProperty(XMLInputFactory.SUPPORT_DTD, Boolean.FALSE);
        factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, Boolean.FALSE);
        return factory;
    }

    @Override
    public void close() throws IOException {
        pkg.close();
    }

    public boolean isDate1904() {
        return date1904;
    }

    public Stream<Sheet> getSheets() {
        return sheets.stream();
    }

    public Optional<Sheet> getSheet(int index) {
        return index < 0 || index >= sheets.size() ? Optional.empty() : Optional.of(sheets.get(index));
    }

    public Sheet getFirstSheet() {
        return sheets.get(0);
    }

    public Optional<Sheet> findSheet(String name) {
        return sheets.stream().filter(sheet -> name.equals(sheet.getName())).findFirst();
    }

    private void readWorkbook(SimpleXmlReader r) throws XMLStreamException {
        while (r.goTo(() -> r.isStartElement("sheets") || r.isStartElement("workbookPr") || r.isEndElement("workbook"))) {
            if ("sheets".equals(r.getLocalName())) {
                r.forEach("sheet", "sheets", this::createSheet);
            } else if ("workbookPr".equals(r.getLocalName())) {
                String date1904Value = r.getAttribute("date1904");
                date1904 = Boolean.parseBoolean(date1904Value);
            } else {
                break;
            }
        }
    }

    private void createSheet(SimpleXmlReader r) {
        String name = r.getAttribute("name");
        String id = r.getAttribute("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        int index = sheets.size();
        sheets.add(new Sheet(this, index, id, name));
    }

    Stream<Row> openStream(Sheet sheet) throws IOException {
        try {
            InputStream inputStream = pkg.getSheetContent(sheet);
            Stream<Row> stream = StreamSupport.stream(new RowSpliterator(this, inputStream), false);
            return stream.onClose(asUncheckedRunnable(inputStream));
        } catch (XMLStreamException e) {
            throw new IOException(e);
        }
    }

    XMLInputFactory getXmlFactory() {
        return factory;
    }

    SST getSharedStringsTable() {
        return sst;
    }

    public static boolean isOOXMLZipHeader(byte[] bytes) {
        return HeaderSignatures.isHeader(bytes, HeaderSignatures.OOXML_FILE_HEADER);
    }

    public static boolean isOLE2Header(byte[] bytes) {
        return HeaderSignatures.isHeader(bytes, HeaderSignatures.OLE_2_SIGNATURE);
    }

    private static Runnable asUncheckedRunnable(Closeable c) {
        return () -> {
            try {
                c.close();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        };
    }

}
