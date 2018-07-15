/**
 * 
 */
package com.github.mforoni.jspreadsheet;

import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.OUTPUT_DIR;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.SHEET1;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import com.github.mforoni.jbasic.io.JFiles;
import com.github.mforoni.jspreadsheet.JxlSpreadsheet;
import jxl.read.biff.BiffException;

/**
 * @author Foroni Marco
 */
public class JxlSpreadsheetTest {

	static final String RATINGS_XLS = "Ratings.xls";
	private static final String CREATE_XLS = "Create.xls";
	private static final Path OUTPUT_FILEPATH = Paths.get(OUTPUT_DIR.toString(), CREATE_XLS);
	private static final String EMPTY_XLS = "Empty.xls";

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#create(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testCreate() throws FileNotFoundException, IOException {
		// try (final JxlSpreadsheet spreadsheet = JxlSpreadsheet.create(JFiles.getResource(CREATE_XLS))) {
		// assertNotNull(spreadsheet);
		// }
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#open(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 * @throws BiffException
	 */
	@Test
	public void testOpen() throws FileNotFoundException, IOException {
		try (final JxlSpreadsheet spreadsheet = JxlSpreadsheet.open(JFiles.fromResource(EMPTY_XLS))) {
			assertNotNull(spreadsheet);
		}
		// try (final JxlSpreadsheet spreadsheet = open(JFiles.getResource(RATINGS_XLS))) {
		// assertNotNull(spreadsheet);
		// }
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#edit(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testEdit() throws FileNotFoundException, IOException {
		// try (final JxlSpreadsheet spreadsheet = edit(JFiles.getResource(RATINGS_XLS))) {
		// assertNotNull(spreadsheet);
		// }
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#getSheetNames()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheetNames() throws FileNotFoundException, IOException {
		try (final JxlSpreadsheet spreadsheet = JxlSpreadsheet.open(JFiles.fromResource(EMPTY_XLS))) {
			final List<String> sheetNames = spreadsheet.getSheetNames();
			assertNotNull(sheetNames);
			assertEquals(1, sheetNames.size());
			assertEquals(SHEET1, sheetNames.get(0));
		}
		// try (final JxlSpreadsheet spreadsheet = JxlSpreadsheet.open(JFiles.getResource(RATINGS_XLS))) {
		// final List<String> sheetNames = spreadsheet.getSheetNames();
		// assertNotNull(sheetNames);
		// assertEquals(EXPECTED_SHEET_NUMBER, sheetNames.size());
		// assertEquals(HEROES, sheetNames.get(0));
		// assertEquals(SHEET1, sheetNames.get(1));
		// }
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#getSheets()}.
	 */
	@Test
	public void testGetSheets() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#getSheet(java.lang.String)}.
	 */
	@Test
	public void testGetSheetString() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#getSheet(int)}.
	 */
	@Test
	public void testGetSheetInt() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#addSheet(java.lang.String)}.
	 */
	@Test
	public void testAddSheet() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#write()}.
	 */
	@Test
	public void testWrite() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.JxlSpreadsheet#close()}.
	 */
	@Test
	public void testClose() {}
}
