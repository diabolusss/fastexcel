package org.dhatim.fastexcel.reader;

import static org.assertj.core.api.Assertions.assertThat;
import static org.dhatim.fastexcel.reader.Resources.open;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Optional;

import org.junit.jupiter.api.Test;

public class HyperlinksTest {

    private static final String ONE_SHEET_2011_XLSX = "/xlsx/hyperlink_stress_test_2011.xlsx";
    private static final String SHEETS_CROSSLINK_NORELS_XLSX = "/xlsx/hyperlink_no_rels_withFormatting_40cells.xlsx";
    private static final String TWO_SHEETS_XLSX = "/xlsx/hyperlink_2sheets_11links.xlsx";
    private static final String LINKS_STRESS_TEST2_XLSX = "/xlsx/hyperlink_stress_test_2023.xlsx";

    private static HashMap<String, String> ONE_SHEET_HYPERLINKS = new  HashMap<String,String>(){
    	private static final long serialVersionUID = -2729555103231479585L;
		{
	    	put("1A1","http://www.sheetjs.com/");
	    	put("1A2","http://oss.sheetjs.com/");
	    	put("1A3","http://oss.sheetjs.com/");//TODO 2023-08-11 #foo anchor part is missing
	    	put("1A4","mailto:dev@sheetjs.com");
	    	put("1A5","mailto:dev@sheetjs.com?subject=hyperlink");
	    	put("1A6","../../sheetjs/Documents/Test.xlsx");
	    	put("1A7","http://sheetjs.com/");
	    }
    };
    private static HashMap<String, String> ONE_SHEET_2011_HYPERLINKS = new  HashMap<String,String>(){
    	private static final long serialVersionUID = -2729555103231479585L;
		{
	    	put("1A1","http://www.sheetjs.com");
	    	put("1A2","http://oss.sheetjs.com");
	    	put("1A3","http://oss.sheetjs.com");//TODO 2023-08-11 #foo anchor part is missing
	    	put("1A4","mailto:dev@sheetjs.com");
	    	put("1A5","mailto:dev@sheetjs.com?subject=hyperlink");
	    	put("1A6","../../sheetjs/Documents/Test.xlsx");
	    	put("1A7","http://sheetjs.com");
	    }
    };
    private static HashMap<String, String> TWO_SHEET_2ND_SHEET_HYPERLINKS = new  HashMap<String,String>(){
    	private static final long serialVersionUID = -2729555103231479585L;
		{
	    	put("2A1","https://example/link1");
	    	put("2B2","https://example/link2");
	    	put("2D4","https://example/link3");
	    	put("2E5","https://example/link4");
	    }
    };

    @Test
    void testWithHyperlinksTwoSheetsCrosslinkNoHyperlinksWorkbookFromInputStream() throws IOException {
    	ReadingOptions options = new ReadingOptions(false, false, true);
        try (InputStream inputStream = open(SHEETS_CROSSLINK_NORELS_XLSX);
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, options)) {
        	assertWithHyperlinksTwoSheetwCrosslinkHyperlinkNoRelsWorkbook(excel);
        }
    }

    @Test
    void testWithHyperlinksOneSheetWorkbookFromInputStream() throws IOException {
        try (InputStream inputStream = open(ONE_SHEET_2011_XLSX);
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(false, false, true))) {
        	assertWithHyperlinksOneSheetWorkbook(excel);
        }
    }

    @Test
    void testWithHyperlinksTwoSheetsWorkbookFromInputStream() throws IOException {
        try (InputStream inputStream = open(TWO_SHEETS_XLSX);
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(false, false, true))) {
        	assertWithHyperlinksTwoSheetsWorkbook(excel);
        }
    }

    @Test
    void testWithHyperlinksStressTestTwoSheetsWorkbookFromInputStream() throws IOException {
        try (InputStream inputStream = open(LINKS_STRESS_TEST2_XLSX);
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(false, false, true))) {
        	assertWithHyperlinksStressTestTwoSheetsWorkbook(excel, true);
        }
    }

    @Test
    void testSkipHyperlinksStressTestTwoSheetsWorkbookFromInputStream() throws IOException {
        try (InputStream inputStream = open(LINKS_STRESS_TEST2_XLSX);
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(false, false, false))) {
        	assertWithHyperlinksStressTestTwoSheetsWorkbook(excel, false);
        }
    }

    private void assertWithHyperlinksTwoSheetwCrosslinkHyperlinkNoRelsWorkbook(ReadableWorkbook excel) throws IOException {
        Optional<Sheet> optSheet = excel.getSheet(0);
        assertThat(optSheet).isPresent();
        Sheet sheet = optSheet.get();

        assertThat(excel.getLinkIdToLinkString()).isEmpty();

        Row[] rows = sheet.openStream().toArray(Row[]::new);

        assertThat(rows).hasSize(40);

        assertThat(excel.getLinkIdToLinkString()).isEmpty();
    }

    private void assertWithHyperlinksOneSheetWorkbook(ReadableWorkbook excel) throws IOException {
        Optional<Sheet> optSheet = excel.getSheet(0);
        assertThat(optSheet).isPresent();
        Sheet sheet = optSheet.get();

        assertThat(excel.getLinkIdToLinkString()).hasSize(7); //links are populated at initial reading

        Row[] rows = sheet.openStream().toArray(Row[]::new);

        assertThat(rows).hasSize(7);

        assertThat(rows[0].getCell(0)).extracting(Cell::getText).isEqualTo("http://www.sheetjs.com");
        assertThat(rows[1].getCell(0)).extracting(Cell::getText).isEqualTo("OSS");
        assertThat(rows[2].getCell(0)).extracting(Cell::getText).isEqualTo("OSS#foo");
        assertThat(rows[3].getCell(0)).extracting(Cell::getText).isEqualTo("dev at sheetjs dot com");
        assertThat(rows[4].getCell(0)).extracting(Cell::getText).isEqualTo("dev at sheetjs.com subject hyperlink");
        assertThat(rows[5].getCell(0)).extracting(Cell::getText).isEqualTo("file://localhost/Users/sheetjs/Documents/Test.xlsx");
        assertThat(rows[6].getCell(0)).extracting(Cell::getText).isEqualTo("http://sheetjs.com screentip foo bar baz");

        assertThat(excel.getLinkIdToLinkString()).containsAllEntriesOf(ONE_SHEET_2011_HYPERLINKS); //link relationships are updated after reading rows (each sheet)

    }


    private void assertWithHyperlinksTwoSheetsWorkbook(ReadableWorkbook excel) throws IOException {
        Optional<Sheet> sheet = excel.getSheet(0);
        assertThat(sheet).isPresent();
        Row[] rows = sheet.get().openStream().toArray(Row[]::new);

        assertThat(excel.getLinkIdToLinkString()).hasSize(11);
//        System.out.println("==== print map ====================================");
//        for(String key: excel.getLinkIdToLinkString().keySet()){
//        	System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
//        }
//
//        System.out.println("===== print by cell value 1 ===================================");
//        for(int i = 0; i < rows.length; i++){
//        	System.out.println(sheet.get().getIndex()+": ("+rows[i].getCell(0).getAddress()+") "+rows[i].getCell(0).getRawValue()+"="+excel.getLinkIdToLinkString().get((sheet.get().getIndex()+1)+rows[i].getCell(0).getAddress().toString()));
//        }

        assertThat(rows).hasSize(7);

        assertThat(rows[0].getCell(0)).extracting(Cell::getText).isEqualTo("http://www.sheetjs.com");
        assertThat(rows[1].getCell(0)).extracting(Cell::getText).isEqualTo("OSS");
        assertThat(rows[2].getCell(0)).extracting(Cell::getText).isEqualTo("OSS#foo");
        assertThat(rows[3].getCell(0)).extracting(Cell::getText).isEqualTo("dev at sheetjs dot com");
        assertThat(rows[4].getCell(0)).extracting(Cell::getText).isEqualTo("dev at sheetjs.com subject hyperlink");
        assertThat(rows[5].getCell(0)).extracting(Cell::getText).isEqualTo("file://localhost/Users/sheetjs/Documents/Test.xlsx");
        assertThat(rows[6].getCell(0)).extracting(Cell::getText).isEqualTo("http://sheetjs.com screentip foo bar baz");

        assertThat(excel.getLinkIdToLinkString()).containsAllEntriesOf(ONE_SHEET_HYPERLINKS);

        sheet = excel.getSheet(1);
        assertThat(sheet).isPresent();
        rows = sheet.get().openStream().toArray(Row[]::new);

//        System.out.println("===== print by cell value 2 ===================================");
//        for(int i = 0; i < rows.length; i++){
//        	System.out.println(sheet.get().getIndex()+": ("+rows[i].getCell(i).getAddress()+") "+rows[i].getCell(i).getRawValue()+"="+excel.getLinkIdToLinkString().get((sheet.get().getIndex()+1)+rows[i].getCell(i).getAddress().toString()));
//        }

        assertThat(rows).hasSize(6);

        assertThat(rows[0].getCell(0)).extracting(Cell::getText).isEqualTo("link1");
        assertThat(rows[1].getCell(1)).extracting(Cell::getText).isEqualTo("link2");
        assertThat(rows[2].getCell(2)).extracting(Cell::getText).isEqualTo("not link");
        assertThat(rows[3].getCell(3)).extracting(Cell::getText).isEqualTo("link3");
        assertThat(rows[4].getCell(4)).extracting(Cell::getText).isEqualTo("link4");
        assertThat(rows[5].getCell(5)).extracting(Cell::getText).isEqualTo("empty link");

        assertThat(excel.getLinkIdToLinkString()).containsAllEntriesOf(TWO_SHEET_2ND_SHEET_HYPERLINKS);
    }


    private void assertWithHyperlinksStressTestTwoSheetsWorkbook(ReadableWorkbook excel, boolean withHyperlinks) throws IOException {
        Optional<Sheet> sheet = excel.getSheet(0);
        assertThat(sheet).isPresent();
        Row[] rows = sheet.get().openStream().toArray(Row[]::new);

//        System.out.println("==== print map ====================================");
//        for(String key: excel.getLinkIdToLinkString().keySet()){
//        	System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
//        }
        if(withHyperlinks){
        	assertThat(excel.getLinkIdToLinkString()).hasSize(424+71 + 1 /*2F22*/);

        	//not hyperlinks
        	assertThat(excel.getLinkIdToLinkString()).doesNotContainEntry("1rId425", "../tables/table1.xml");
        	assertThat(excel.getLinkIdToLinkString()).doesNotContainEntry("2rId73", "../tables/table2.xml");
        }else{
        	assertThat(excel.getLinkIdToLinkString()).isEmpty();
        }

//        System.out.println("===== print by cell value 1 ===================================");
//        for(int i = 0; i < rows.length; i++){
//        	System.out.println(sheet.get().getIndex()+": ("+rows[i].getCell(0).getAddress()+") "+rows[i].getCell(0).getRawValue()+"="+excel.getLinkIdToLinkString().get((sheet.get().getIndex()+1)+rows[i].getCell(0).getAddress().toString()));
//        }

        assertThat(rows).hasSize(1+424+9);// header + contents + footer

        if(withHyperlinks){
        	assertThat(excel.getLinkIdToLinkString()).containsEntry("1A386", "https://www.gov.pl/attachment/d87b02f0-14ab-473c-a211-efc6db43c96e");
        }

        sheet = excel.getSheet(1);
        assertThat(sheet).isPresent();
        rows = sheet.get().openStream().toArray(Row[]::new);

//        System.out.println("===== print by cell value 2 ===================================");
//        for(int i = 0; i < rows.length; i++){
//        	System.out.println(sheet.get().getIndex()+": ("+rows[i].getCell(0).getAddress()+") "+rows[i].getCell(0).getRawValue()+"="+excel.getLinkIdToLinkString().get((sheet.get().getIndex()+1)+rows[i].getCell(0).getAddress().toString()));
//        }

        if(withHyperlinks){
        	assertThat(rows).hasSize(1+71+9);// header + contents + footer

//        System.out.println("==== print map ====================================");
        	int countA = 0;int countB = 0;int countC = 0;int countD = 0;int countE = 0; int countF = 0;
        	for(String key: excel.getLinkIdToLinkString().keySet()){

        		if(key.contains("A")){
//        		System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
        			countA++;
        		}
        		if(key.contains("B")){
//        		System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
        			countB++;
        		}
        		if(key.contains("C")){
//        		System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
        			countC++;
        		}
        		if(key.contains("D")){
//        		System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
        			countD++;
        		}
        		if(key.contains("E")){
//        		System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
        			countE++;
        		}
        		if(key.contains("F")){
//        		System.out.println(key+"="+excel.getLinkIdToLinkString().get(key));
        			countF++;
        		}
        	}

        	System.out.println("==== print totals ====================================");
        	System.out.println("total A rows have links => "+countA);
        	System.out.println("total B rows have links => "+countB);
        	System.out.println("total C rows have links => "+countC);
        	System.out.println("total D rows have links => "+countD);
        	System.out.println("total E rows have links => "+countE);
        	System.out.println("total F rows have links => "+countF);

        	assertThat(excel.getLinkIdToLinkString()).containsEntry("2A56", "https://www.gov.pl/attachment/c5cdd582-4e0b-4737-8d70-4ea14849c604");
        }else{
        	assertThat(excel.getLinkIdToLinkString()).isEmpty();
        }
    }

}
