/**
 * 
 */
package com.github.mforoni.jspreadsheet;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import com.github.mforoni.jbasic.io.JFiles;
import com.github.mforoni.jspreadsheet.OdsSpreadsheet;
import com.github.mforoni.jspreadsheet.Sheet;

/**
 * @author Foroni Marco
 */
public class OdsSpreadsheetTest {

	static final String RATINGS_ODS = "Ratings.ods";
	static final String HEROES = "Heroes";
	static final String SHEET1 = "Sheet1";
	static final String SHEET2 = "Sheet2";
	static final int EXPECTED_SHEET_NUMBER = 2;
	static final String EXPECTED_MSG_INDEX_OUT_OF_RANGE = "Sheet index (2) is out of range (0..1)";
	static final Path OUTPUT_DIR = Paths.get("output", "test");
	private static final String CREATE_ODS = "Create.ods";
	private static final Path OUTPUT_FILEPATH = Paths.get(OUTPUT_DIR.toString(), CREATE_ODS);

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
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#create(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testCreate() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.create(JFiles.fromResource(RATINGS_ODS))) {
			assertNotNull(spreadsheet);
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#edit(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testEdit() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.edit(JFiles.fromResource(RATINGS_ODS))) {
			assertNotNull(spreadsheet);
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#open(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testOpen() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
			assertNotNull(spreadsheet);
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#getSheetNames()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheetNames() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
			final List<String> sheetNames = spreadsheet.getSheetNames();
			assertNotNull(sheetNames);
			assertEquals(EXPECTED_SHEET_NUMBER, sheetNames.size());
			assertEquals(HEROES, sheetNames.get(0));
			assertEquals(SHEET1, sheetNames.get(1));
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#getSheets()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheets() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
			final List<Sheet> sheets = spreadsheet.getSheets();
			assertNotNull(sheets);
			assertEquals(EXPECTED_SHEET_NUMBER, sheets.size());
			assertNotNull(sheets.get(0));
			assertNotNull(sheets.get(1));
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#getSheet(java.lang.String)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheetString() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
			assertNotNull(spreadsheet.getSheet(HEROES));
			assertNotNull(spreadsheet.getSheet(SHEET1));
			try {
				spreadsheet.getSheet(SHEET2);
				fail("Exception not thrown");
			} catch (final IllegalArgumentException e) {
				assertNotNull(e);
				assertEquals("Cannot find sheet having name " + SHEET2, e.getMessage());
			}
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#getSheet(int)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheetInt() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
			assertNotNull(spreadsheet.getSheet(0));
			assertNotNull(spreadsheet.getSheet(1));
			try {
				spreadsheet.getSheet(2);
				fail("Exception not thrown");
			} catch (final IllegalArgumentException e) {
				assertNotNull(e);
				assertEquals(EXPECTED_MSG_INDEX_OUT_OF_RANGE, e.getMessage());
			}
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#addSheet(java.lang.String)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testAddSheet() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.create(OUTPUT_FILEPATH.toFile())) {
			final Sheet sheet = spreadsheet.addSheet(SHEET1);
			assertNotNull(sheet);
			assertEquals(SHEET1, sheet.getName());
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#write()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testWrite() throws FileNotFoundException, IOException {
		try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.create(OUTPUT_FILEPATH.toFile())) {
			spreadsheet.addSheet(SHEET1);
			spreadsheet.write();
		}
		assertTrue(OUTPUT_FILEPATH.toFile().exists());
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.OdsSpreadsheet#close()}.
	 */
	@Test
	public void testClose() {}
}
