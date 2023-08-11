package org.dhatim.fastexcel.reader;

import static java.lang.String.format;
import static org.dhatim.fastexcel.reader.DefaultXMLInputFactory.factory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.regex.Pattern;

import javax.xml.stream.XMLStreamException;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.commons.compress.utils.SeekableInMemoryByteChannel;

class OPCPackage implements AutoCloseable {
    private static final Pattern filenameRegex = Pattern.compile("^(.*/)([^/]+)$");
    private final ZipFile zip;
    private final Map<String, String> workbookPartsById;
    private final PartEntryNames parts;
    private final List<String> formatIdList;
    private Map<String, String> fmtIdToFmtString;
    /**
     * if withHyperlinks, then:
     *  - try to read hyperlink ids from /xl/worksheet/_rels/.rels (first pass)
     *  - after processing rows, try to read <hyperlinks> and refresh map key to point to {@link CellAddress} (second pass)
     */
    private Map<String, String> linkIdToLinkString;

    private OPCPackage(File zipFile) throws IOException {
        this(zipFile, false, false);
    }

    private OPCPackage(File zipFile, boolean withFormat, boolean withHyperlinks) throws IOException {
        this(new ZipFile(zipFile), withFormat, withHyperlinks);
    }

    private OPCPackage(SeekableInMemoryByteChannel channel, boolean withStyle, boolean withHyperlinks) throws IOException {
        this(new ZipFile(channel), withStyle, withHyperlinks);
    }

    private OPCPackage(ZipFile zip, boolean withFormat, boolean withHyperlinks) throws IOException {
        try {
            this.zip = zip;
            this.parts = extractPartEntriesFromContentTypes(withHyperlinks);
            if (withFormat) {
                this.formatIdList = extractFormat(parts.style);
            } else {
                this.formatIdList = Collections.emptyList();
            }

            //read complete lists of hyperlinks (first pass)
            //reading links from sheet will not happen in case of empty map (second pass)
            this.linkIdToLinkString = withHyperlinks && parts._relsCount > 0 ? readWorkbookHyperlinkIds(PartEntryNames.SHEET_HYPERLINKS_RELS_FILE) : Collections.emptyMap();

            this.workbookPartsById = readWorkbookPartsIds(relsNameFor(parts.workbook));
        } catch (XMLStreamException e) {
            throw new IOException(e);
        }
    }

    private static String relsNameFor(String entryName) {
        return filenameRegex.matcher(entryName).replaceFirst("$1_rels/$2.rels");
    }

    private Map<String, String> readWorkbookPartsIds(String workbookRelsEntryName) throws IOException, XMLStreamException {
        String xlFolder = workbookRelsEntryName.substring(0, workbookRelsEntryName.indexOf("_rel"));
        Map<String, String> partsIdById = new HashMap<>();
        SimpleXmlReader rels = new SimpleXmlReader(factory, getRequiredEntryContent(workbookRelsEntryName));
        while (rels.goTo("Relationship")) {
            String id = rels.getAttribute("Id");
            String target = rels.getAttribute("Target");
            // if name does not start with /, it is a relative path
            if (!target.startsWith("/")) {
                target = xlFolder + target;
            } // else it is an absolute path
            partsIdById.put(id, target);
        }
        return partsIdById;
    }

    /**
     * Try to read hyperlink relationships.
     *  On success map is not yet usable, because {@link CellAddress}'es may be read only from the end of the sheet.
     */
    private Map<String, String> readWorkbookHyperlinkIds(String workbookRelsEntryName) throws IOException, XMLStreamException {
        if (parts._relsCount <= 0) { return Collections.emptyMap(); }

        Map<String, String> hyperlinksById = new HashMap<>();
        for(int i = 1; i <= parts._relsCount; i++) {
        	String wbSheetRelsEntryName = String.format(workbookRelsEntryName, i);
        	try(SimpleXmlReader rels = new SimpleXmlReader(factory, getRequiredEntryContent(wbSheetRelsEntryName))){
	         	while (rels.goTo("Relationship")){
	        		String id = rels.getAttribute("Id");
	        		String type = rels.getAttribute("Type");
	        		String target = rels.getAttribute("Target");
	        		//String targetMode = rels.getAttribute("TargetMode");

	        		//process only hyperlinks
	        		//if(!"External".equals(targetMode)){ continue; }
	        		if(type == null || !type.endsWith("/hyperlink")){ continue; }

	        		hyperlinksById.put(i+id, target);
	        	}
	        }
        }
        return hyperlinksById;
    }

    /**
     * @param withHyperlinks - when set, try to estimate possible _rels file count (based on worksheet count, because _rels may not be listed)
     * @return
     * @throws XMLStreamException
     * @throws IOException
     */
    private PartEntryNames extractPartEntriesFromContentTypes(boolean  withHyperlinks) throws XMLStreamException, IOException {
        PartEntryNames entries = new PartEntryNames();
        final String contentTypesXml = "[Content_Types].xml";
        try (SimpleXmlReader reader = new SimpleXmlReader(factory, getRequiredEntryContent(contentTypesXml))) {
            while (reader.goTo(() -> reader.isStartElement("Override"))) {
                String contentType = reader.getAttributeRequired("ContentType");
                if (PartEntryNames.WORKBOOK_MAIN_CONTENT_TYPE.equals(contentType)
                        || PartEntryNames.WORKBOOK_EXCEL_MACRO_ENABLED_MAIN_CONTENT_TYPE.equals(contentType)) {
                    entries.workbook = reader.getAttributeRequired("PartName");
                } else if (PartEntryNames.SHARED_STRINGS_CONTENT_TYPE.equals(contentType)) {
                    entries.sharedStrings = reader.getAttributeRequired("PartName");
                } else if (PartEntryNames.STYLE_CONTENT_TYPE.equals(contentType)) {
                    entries.style = reader.getAttributeRequired("PartName");
                } else if (withHyperlinks && PartEntryNames.WORKBOOK_WORKSHEET_TYPE.equals(contentType)){
                	String partName = reader.getAttribute("PartName");
                	if(partName.startsWith("/xl/worksheets/sheet")){ //count by sheets as _rels may be missing
                		entries._relsCount++;
                	}
                }
                if (entries.isFullyFilled() && !withHyperlinks) {
                    break;
                }
            }
            if (entries.workbook == null) {
                // in case of a default workbook path, we got this
                // <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
                entries.workbook = "/xl/workbook.xml";
            }
        }
        return entries;
    }

    private List<String> extractFormat(String styleXml) throws XMLStreamException, IOException {
        List<String> fmtIdList = new ArrayList<>();
        fmtIdToFmtString = new HashMap<>();
        try (SimpleXmlReader reader = new SimpleXmlReader(factory, getRequiredEntryContent(styleXml))) {
            AtomicBoolean insideCellXfs = new AtomicBoolean(false);
            while (reader.goTo(() -> reader.isStartElement("numFmt") ||
                reader.isStartElement("cellXfs") || reader.isEndElement("cellXfs") ||
                insideCellXfs.get())) {
                if (reader.isStartElement("cellXfs")) {
                    insideCellXfs.set(true);
                } else if (reader.isEndElement("cellXfs")) {
                    insideCellXfs.set(false);
                }
                if ("numFmt".equals(reader.getLocalName())) {
                    String formatCode = reader.getAttributeRequired("formatCode");
                    fmtIdToFmtString.put(reader.getAttributeRequired("numFmtId"), formatCode);
                } else if (insideCellXfs.get() && reader.isStartElement("xf")) {
                    fmtIdList.add(reader.getAttribute ("numFmtId"));
                }
            }
        }
        return fmtIdList;
    }

    private InputStream getRequiredEntryContent(String name) throws IOException {
        return Optional.ofNullable(getEntryContent(name))
            .orElseThrow(() -> new ExcelReaderException(name + " not found"));
    }

    static OPCPackage open(File inputFile) throws IOException {
        return open(inputFile, false, false);
    }

    static OPCPackage open(File inputFile, boolean withFormat, boolean withHyperlinks) throws IOException {
        return new OPCPackage(inputFile, withFormat, withHyperlinks);
    }

    static OPCPackage open(InputStream inputStream) throws IOException {
        return open(inputStream, false, false);
    }

    static OPCPackage open(InputStream inputStream, boolean withFormat, boolean withHyperlinks) throws IOException {
        byte[] compressedBytes = IOUtils.toByteArray(inputStream);
        return new OPCPackage(new SeekableInMemoryByteChannel(compressedBytes), withFormat, withHyperlinks);
    }

    InputStream getSharedStrings() throws IOException {
        return getEntryContent(parts.sharedStrings);
    }

    private ZipArchiveEntry getEntry(String name) throws IOException {
    	if (name == null) { return null; }
        if (name.startsWith("/")) { name = name.substring(1); }

        final ZipArchiveEntry entry = zip.getEntry(name);
        if (entry != null) { return entry; }

        // to be case insensitive
        Enumeration<ZipArchiveEntry> entries = zip.getEntries();
        while (entries.hasMoreElements()) {
            ZipArchiveEntry e = entries.nextElement();
            if (e.getName().equalsIgnoreCase(name)) {
                return e;
            }
        }
        return null;
    }

    private InputStream getEntryContent(String name) throws IOException {
        ZipArchiveEntry entry = getEntry(name);
        return entry == null ? null : zip.getInputStream(entry);
    }

    @Override
    public void close() throws IOException {
        zip.close();
    }

    public InputStream getWorkbookContent() throws IOException {
        return getRequiredEntryContent(parts.workbook);
    }

    public InputStream getSheetContent(Sheet sheet) throws IOException {
        String name = this.workbookPartsById.get(sheet.getId());
        if (name == null) {
            String msg = format("Sheet#%s '%s' is missing an entry in workbook rels (for id: '%s')",
                sheet.getIndex(), sheet.getName(), sheet.getId());
            throw new ExcelReaderException(msg);
        }
        return getRequiredEntryContent(name);
    }

    public List<String> getFormatList() {
        return formatIdList;
    }

    public Map<String, String> getFmtIdToFmtString() {
        return fmtIdToFmtString;
    }

    /** <pre>
     * @return empty if not withHyperlinks, otherwise:
     *  - mapped hyperlink ids from /xl/worksheet/_rels/.rels (first pass => {@link OPCPackage#open(InputStream, boolean, boolean)})
     *  - updated map where key points to {@link CellAddress} (second pass => {@link RowSpliterator#nextHyperlink()}
     */
    public Map<String, String> getLinkIdToLinkString() {
		return linkIdToLinkString;
	}

	private static class PartEntryNames {
        public static final String WORKBOOK_MAIN_CONTENT_TYPE =
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        public static final String WORKBOOK_EXCEL_MACRO_ENABLED_MAIN_CONTENT_TYPE =
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
        public static final String SHARED_STRINGS_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        public static final String STYLE_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        public static final String WORKBOOK_WORKSHEET_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        public static final String SHEET_HYPERLINKS_RELS_FILE = "/xl/worksheets/_rels/sheet%d.xml.rels";
        String workbook;
        String sharedStrings;
        String style;
        int _relsCount;

        boolean isFullyFilled() {
            return workbook != null && sharedStrings != null && style != null;
        }
    }
}
