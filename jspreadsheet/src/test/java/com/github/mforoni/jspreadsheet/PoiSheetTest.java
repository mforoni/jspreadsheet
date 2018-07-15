/**
 * 
 */
package com.github.mforoni.jspreadsheet;

import static com.github.mforoni.jspreadsheet.OdsSheetTest.HEROES_EXPECTED_COLUMNS;
import static com.github.mforoni.jspreadsheet.OdsSheetTest.HEROES_EXPECTED_ROWS;
import static com.github.mforoni.jspreadsheet.OdsSheetTest.SHEET1_EXPECTED_LAST_ROW;
import static com.github.mforoni.jspreadsheet.OdsSheetTest.SHEET1_EXPECTED_ROWS;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.HEROES;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.SHEET1;
import static com.github.mforoni.jspreadsheet.PoiSpreadsheetTest.RATINGS_XLSX;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.joda.time.LocalDate;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import com.github.mforoni.jbasic.io.JFiles;
import com.github.mforoni.jspreadsheet.PoiSpreadsheet;
import com.github.mforoni.jspreadsheet.Sheet;

/**
 * @author Foroni Marco
 */
public class PoiSheetTest {

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
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.PoiSheet#PoiSheet(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Workbook, boolean)}.
	 */
	@Test
	public void testPoiSheet() {}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getName()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetName() throws IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			Sheet sheet = spreadsheet.getSheet(HEROES);
			assertNotNull(sheet);
			assertEquals(HEROES, sheet.getName());
			sheet = spreadsheet.getSheet(SHEET1);
			assertNotNull(sheet);
			assertEquals(SHEET1, sheet.getName());
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getRows()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetRows() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			Sheet sheet = spreadsheet.getSheet(HEROES);
			assertEquals(HEROES_EXPECTED_ROWS, sheet.getRows());
			sheet = spreadsheet.getSheet(SHEET1);
			assertEquals(SHEET1_EXPECTED_ROWS, sheet.getRows());
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getLastRow()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetLastRow() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			Sheet sheet = spreadsheet.getSheet(HEROES);
			assertEquals(HEROES_EXPECTED_ROWS, sheet.getLastRow());
			sheet = spreadsheet.getSheet(SHEET1);
			assertEquals(SHEET1_EXPECTED_LAST_ROW, sheet.getLastRow());
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getColumns()}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetColumns() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			Sheet sheet = spreadsheet.getSheet(HEROES);
			assertEquals(HEROES_EXPECTED_COLUMNS, sheet.getColumns());
			sheet = spreadsheet.getSheet(SHEET1);
			assertEquals(0, sheet.getColumns());
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getLastColumn(int)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetLastColumn() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			Sheet sheet = spreadsheet.getSheet(HEROES);
			assertEquals(HEROES_EXPECTED_COLUMNS, sheet.getLastColumn(0));
			assertEquals(1, sheet.getLastColumn(5));
			sheet = spreadsheet.getSheet(SHEET1);
			assertEquals(0, sheet.getLastColumn(0));
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getRawValue(int, int)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetRawValue() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			final Sheet sheet = spreadsheet.getSheet(HEROES);
			// row 1
			assertEquals("Abel", sheet.getRawValue(1, 0));
			assertEquals("39", sheet.getRawValue(1, 1));
			assertEquals("SUM(B2:F2)", sheet.getRawValue(1, 6));
			assertEquals("3.5", sheet.getRawValue(1, 7));
			assertEquals("2-Feb-17", sheet.getRawValue(1, 8));
			assertEquals("true", sheet.getRawValue(1, 9));
			// row 5
			assertEquals("Anna", sheet.getRawValue(5, 0));
			assertEquals("", sheet.getRawValue(5, 1));
			assertEquals("", sheet.getRawValue(5, 6));
			assertEquals("", sheet.getRawValue(5, 7));
			assertEquals("", sheet.getRawValue(5, 8));
			assertEquals("", sheet.getRawValue(5, 9));
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getObject(int, int)}.
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	@Test
	public void testGetObject() throws FileNotFoundException, IOException {
		try (final PoiSpreadsheet spreadsheet = PoiSpreadsheet.open(JFiles.fromResource(RATINGS_XLSX))) {
			final Sheet sheet = spreadsheet.getSheet(HEROES);
			// row 1
			assertEquals("Abel", sheet.getObject(1, 0));
			assertEquals(39D, sheet.getObject(1, 1));
			assertEquals(154D, sheet.getObject(1, 6));
			assertEquals(3.5D, sheet.getObject(1, 7));
			assertEquals(new LocalDate(2017, 2, 2).toDate(), sheet.getObject(1, 8));
			assertEquals(true, sheet.getObject(1, 9));
			// row 5
			assertEquals("Anna", sheet.getObject(5, 0));
			assertNull(sheet.getObject(5, 1));
			assertNull(sheet.getObject(5, 6));
			assertNull(sheet.getObject(5, 7));
			assertNull(sheet.getObject(5, 8));
			assertNull(sheet.getObject(5, 9));
		}
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getString(int, int)}.
	 */
	@Test
	public void testGetString() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getDouble(int, int)}.
	 */
	@Test
	public void testGetDouble() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getDate(int, int)}.
	 */
	@Test
	public void testGetDate() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#getBoolean(int, int)}.
	 */
	@Test
	public void testGetBoolean() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#getRow(int)}.
	 */
	@Test
	public void testGetRowInt() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#getRow(int, int, int)}.
	 */
	@Test
	public void testGetRowIntIntInt() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#hideColumn(int)}.
	 */
	@Test
	public void testHideColumn() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setAutoSize()}.
	 */
	@Test
	public void testSetAutoSize() {}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.PoiSheet#setAutoSize(int, int)}.
	 */
	@Test
	public void testSetAutoSizeIntInt() {
		// TODO
	}

	/**
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.PoiSheet#setString(int, int, java.lang.String, com.github.mforoni.jspreadsheet.SSCellFormat)}.
	 */
	@Test
	public void testSetStringIntIntStringSSCellFormat() {
		// TODO
	}

	/**
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.PoiSheet#setDouble(int, int, java.lang.Double, com.github.mforoni.jspreadsheet.SSCellFormat)}.
	 */
	@Test
	public void testSetDoubleIntIntDoubleSSCellFormat() {
		// TODO
	}

	/**
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.PoiSheet#setDate(int, int, java.util.Date, com.github.mforoni.jspreadsheet.SSCellFormat)}.
	 */
	@Test
	public void testSetDateIntIntDateSSCellFormat() {
		// TODO
	}

	/**
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.PoiSheet#setBoolean(int, int, java.lang.Boolean, com.github.mforoni.jspreadsheet.SSCellFormat)}.
	 */
	@Test
	public void testSetBooleanIntIntBooleanSSCellFormat() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setObject(int, int, java.lang.Object)}.
	 */
	@Test
	public void testSetObjectIntIntObject() {
		// TODO
	}

	/**
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setObject(int, int, java.lang.Object, com.github.mforoni.jspreadsheet.SSCellFormat)}.
	 */
	@Test
	public void testSetObjectIntIntObjectSSCellFormat() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setString(int, int, java.lang.String)}.
	 */
	@Test
	public void testSetStringIntIntString() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setDouble(int, int, java.lang.Double)}.
	 */
	@Test
	public void testSetDoubleIntIntDouble() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setDate(int, int, java.util.Date)}.
	 */
	@Test
	public void testSetDateIntIntDate() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setBoolean(int, int, java.lang.Boolean)}.
	 */
	@Test
	public void testSetBooleanIntIntBoolean() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#add(int, int, java.lang.Double)}.
	 */
	@Test
	public void testAdd() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, java.lang.Object[])}.
	 */
	@Test
	public void testSetRowIntObjectArray() {
		// TODO
	}

	/**
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, java.lang.Object[], com.github.mforoni.jspreadsheet.SSCellFormat)}.
	 */
	@Test
	public void testSetRowIntObjectArraySSCellFormat() {
		// TODO
	}

	/**
	 * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, int, java.lang.Object[])}.
	 */
	@Test
	public void testSetRowIntIntObjectArray() {
		// TODO
	}

	/**
	 * Test method for
	 * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, int, java.lang.Object[], com.github.mforoni.jspreadsheet.SSCellFormat)}.
	 */
	@Test
	public void testSetRowIntIntObjectArraySSCellFormat() {
		// TODO
	}
}
