/**
 * 
 */
package com.github.mforoni.jspreadsheet;

import static com.github.mforoni.jspreadsheet.JxlSpreadsheetTest.RATINGS_XLS;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.EXPECTED_MSG_INDEX_OUT_OF_RANGE;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.EXPECTED_SHEET_NUMBER;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.HEROES;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.OUTPUT_DIR;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.SHEET1;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.SHEET2;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import com.github.mforoni.jbasic.io.JFiles;
import com.github.mforoni.jspreadsheet.PoiSpreadsheet;
import com.github.mforoni.jspreadsheet.Sheet;

/**
 * @author Foroni Marco
 */
public class PoiSpreadsheetTest {

	static final String RATINGS_XLSX = "Ratings.xlsx";
	private static final String CREATE_XLSX = "Create.xlsx";
	private static final Path CREATE_XLSX_PATH = Paths.get(OUTPUT_DIR.toString(), CREATE_XLSX);

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
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#create(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testCreate() throws FileNotFoundException, IOException {
    Files.deleteIfExists(CREATE_XLSX_PATH);
    try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.create(CREATE_XLSX_PATH.toFile())) {
      assertNotNull(spreadsheet);
    }
		com.google.common.io.Files.createParentDirs(CREATE_XLSX_PATH.toFile());
    Files.createFile(CREATE_XLSX_PATH);
    try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.create(CREATE_XLSX_PATH.toFile())) {
      fail("Exception is expected");
    } catch(IllegalStateException e) {
      final String expectedMsg = String.format("Cannot create file %s: the file already exist", CREATE_XLSX_PATH.toFile());
      assertEquals(expectedMsg, e.getMessage());
    }
  }

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#edit(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testEdit() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.edit(JFiles.fromResource(RATINGS_XLSX))) {
			assertNotNull(spreadsheet);
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#open(java.io.File)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testOpen() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			assertNotNull(spreadsheet);
		}
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLS))) {
			assertNotNull(spreadsheet);
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#getSheetNames()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheetNames() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			final List<String> sheetNames = spreadsheet.getSheetNames();
			assertNotNull(sheetNames);
			assertEquals(EXPECTED_SHEET_NUMBER, sheetNames.size());
			assertEquals(HEROES, sheetNames.get(0));
			assertEquals(SHEET1, sheetNames.get(1));
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#getSheets()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheets() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			final List<Sheet> sheets = spreadsheet.getSheets();
			assertNotNull(sheets);
			assertEquals(EXPECTED_SHEET_NUMBER, sheets.size());
			assertNotNull(sheets.get(0));
			assertNotNull(sheets.get(1));
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#getSheet(java.lang.String)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheetString() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
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
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#getSheet(int)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetSheetInt() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
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
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#addSheet(java.lang.String)}.
	 * 
	 * @throws IOException
	 */
	@Test
	public void testAddSheet() throws IOException {
    Files.deleteIfExists(CREATE_XLSX_PATH);
    try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.create(CREATE_XLSX_PATH.toFile())) {
      spreadsheet.addSheet(SHEET1);
      spreadsheet.write();
    }
    assertTrue(CREATE_XLSX_PATH.toFile().exists());
    try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(CREATE_XLSX_PATH.toFile())) {
      final Sheet sheet = spreadsheet.getSheet(SHEET1);
      assertNotNull(sheet);
      assertEquals(SHEET1, sheet.getName());
    }
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#write()}.
	 * 
	 * @throws IOException
	 */
	@Test
	public void testWrite() throws IOException {
	  // same test of testAddSheet()
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSpreadsheet#close()}.
	 */
	@Test
	public void testClose() {}
}
