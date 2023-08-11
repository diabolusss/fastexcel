package org.dhatim.fastexcel.reader;

public class ReadingOptions {
    public static final ReadingOptions DEFAULT_READING_OPTIONS = new ReadingOptions(false, false, false);
    private final boolean withCellFormat;
    private final boolean cellInErrorIfParseError;
    private final boolean withHyperlinks;

    /**
     * @param withCellFormat          If true, extract cell formatting
     * @param cellInErrorIfParseError If true, cell type is ERROR if it is not possible to parse cell value.
     *                                If false, an exception is throw when there is a parsing error
     * @param withHyperlinks		  If true, try to exctract sheet cell /hyperlink relationships
     */
    public ReadingOptions(boolean withCellFormat, boolean cellInErrorIfParseError, boolean withHyperlinks) {
        this.withCellFormat = withCellFormat;
        this.cellInErrorIfParseError = cellInErrorIfParseError;
        this.withHyperlinks = withHyperlinks;
    }

    public ReadingOptions(boolean withCellFormat, boolean cellInErrorIfParseError){
    	this(withCellFormat, cellInErrorIfParseError, false);
    }

    /**
     * @return true to try to extract cell hyperlinks
     */
    public boolean isWithHyperlinks() {
        return withHyperlinks;
    }

    /**
     * @return true for extract cell formatting
     */
    public boolean isWithCellFormat() {
        return withCellFormat;
    }

    /**
     * @return true for cell type is ERROR if it is not possible to parse cell value,
     * false for an exception is throw when there is a parsing error
     */
    public boolean isCellInErrorIfParseError() {
        return cellInErrorIfParseError;
    }
}
