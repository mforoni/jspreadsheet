/**
 * 
 */
package com.github.mforoni.jspreadsheet;

import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.HEROES;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.RATINGS_ODS;
import static com.github.mforoni.jspreadsheet.OdsSpreadsheetTest.SHEET1;
import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import org.joda.time.LocalDate;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import com.github.mforoni.jbasic.io.JFiles;

/**
 * @author Foroni Marco
 */
public class OdsSheetTest {
  static final int HEROES_EXPECTED_ROWS = 10;
  static final int HEROES_EXPECTED_COLUMNS = 10;
  static final int SHEET1_EXPECTED_ROWS = 1; // 0 is not possible: at least one row must be defined
  static final int SHEET1_EXPECTED_LAST_ROW = 0;
  static final int SHEET1_EXPECTED_COLUMNS = 1; // 0 is not possible: at least one column must be
                                                // defined
  static final int SHEET1_EXPECTED_LAST_COLUMN = 0;
  private static final Object[] EXPECTED_ROW_1 = {"Abel", new BigDecimal(39), new BigDecimal(33),
      new BigDecimal(32), new BigDecimal(25), new BigDecimal(25), new BigDecimal(154),
      new BigDecimal(3.5), new LocalDate(2017, 2, 2).toDate(), true};
  private static final Object[] EXPECTED_ROW_5 =
      {"Anna", null, null, null, null, null, null, null, null, null};

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
   * {@link com.github.mforoni.jspreadsheet.OdsSheet#OdsSheet(org.jopendocument.dom.spreadsheet.Sheet)}.
   */
  @Test
  public void testOdsSheet() {}

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getName()}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetName() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      Sheet sheet = spreadsheet.getSheet(HEROES);
      assertNotNull(sheet);
      assertEquals(HEROES, sheet.getName());
      sheet = spreadsheet.getSheet(SHEET1);
      assertNotNull(sheet);
      assertEquals(SHEET1, sheet.getName());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getRows()}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetRows() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      Sheet sheet = spreadsheet.getSheet(HEROES);
      assertEquals(HEROES_EXPECTED_ROWS, sheet.getRows());
      sheet = spreadsheet.getSheet(SHEET1);
      assertEquals(SHEET1_EXPECTED_ROWS, sheet.getRows());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getLastRow()}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetLastRow() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      Sheet sheet = spreadsheet.getSheet(HEROES);
      assertEquals(HEROES_EXPECTED_ROWS, sheet.getLastRow());
      sheet = spreadsheet.getSheet(SHEET1);
      assertEquals(SHEET1_EXPECTED_LAST_ROW, sheet.getLastRow());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getColumns()}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetColumns() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      Sheet sheet = spreadsheet.getSheet(HEROES);
      assertEquals(HEROES_EXPECTED_COLUMNS, sheet.getColumns());
      sheet = spreadsheet.getSheet(SHEET1);
      assertEquals(SHEET1_EXPECTED_COLUMNS, sheet.getColumns());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getLastColumn(int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetLastColumn() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      Sheet sheet = spreadsheet.getSheet(HEROES);
      assertEquals(HEROES_EXPECTED_COLUMNS, sheet.getLastColumn(0));
      assertEquals(1, sheet.getLastColumn(5));
      sheet = spreadsheet.getSheet(SHEET1);
      assertEquals(0, sheet.getLastColumn(0));
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getRawValue(int, int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetRawValue() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      // row 1
      assertEquals("Abel", sheet.getRawValue(1, 0));
      assertEquals("39", sheet.getRawValue(1, 1));
      assertEquals("of:=SUM([.B2:.F2])", sheet.getRawValue(1, 6));
      assertEquals("3.5", sheet.getRawValue(1, 7));
      assertEquals("02/02/17", sheet.getRawValue(1, 8));
      assertEquals("TRUE", sheet.getRawValue(1, 9));
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
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getObject(int, int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetObject() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      // row 1
      assertEquals("Abel", sheet.getObject(1, 0));
      assertEquals(new BigDecimal(39), sheet.getObject(1, 1));
      assertEquals(new BigDecimal(154), sheet.getObject(1, 6));
      assertEquals(new BigDecimal(3.5), sheet.getObject(1, 7));
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
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getString(int, int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetString() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      // row 1
      assertEquals("Abel", sheet.getString(1, 0));
      assertGetStringThrowsException(sheet, 1, 1);
      assertGetStringThrowsException(sheet, 1, 6);
      assertGetStringThrowsException(sheet, 1, 7);
      assertGetStringThrowsException(sheet, 1, 8);
      assertGetStringThrowsException(sheet, 1, 9);
      // row 5
      assertEquals("Anna", sheet.getString(5, 0));
      assertNull(sheet.getString(5, 1));
      assertNull(sheet.getString(5, 6));
      assertNull(sheet.getString(5, 7));
      assertNull(sheet.getString(5, 8));
      assertNull(sheet.getString(5, 9));
    }
  }

  private static void assertGetStringThrowsException(final Sheet sheet, final int row,
      final int col) {
    try {
      sheet.getString(row, col);
      fail("Exception not thrown");
    } catch (final IllegalStateException e) {
      assertEquals("Cannot retrieve a String from a cell not having type STRING", e.getMessage());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getDouble(int, int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetDouble() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      // row 1
      assertGetDoubleThrowsException(sheet, 1, 0);
      assertEquals(new Double(39), sheet.getDouble(1, 1));
      assertEquals(new Double(154), sheet.getDouble(1, 6));
      assertEquals(new Double(3.5), sheet.getDouble(1, 7));
      assertGetDoubleThrowsException(sheet, 1, 8);
      assertGetDoubleThrowsException(sheet, 1, 9);
      // row 5
      assertGetDoubleThrowsException(sheet, 5, 0);
      assertNull(sheet.getDouble(5, 1));
      assertNull(sheet.getDouble(5, 6));
      assertNull(sheet.getDouble(5, 7));
      assertNull(sheet.getDouble(5, 8));
      assertNull(sheet.getDouble(5, 9));
    }
  }

  private static void assertGetDoubleThrowsException(final Sheet sheet, final int row,
      final int col) {
    try {
      sheet.getDouble(row, col);
      fail("Exception not thrown");
    } catch (final IllegalStateException e) {
      assertEquals(
          "Cannot retrieve a Double from a cell not having type FLOAT, CURRENCY or PERCENTAGE",
          e.getMessage());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getDate(int, int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetDate() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      // row 1
      assertGetDateThrowsException(sheet, 1, 0);
      assertGetDateThrowsException(sheet, 1, 1);
      assertGetDateThrowsException(sheet, 1, 6);
      assertGetDateThrowsException(sheet, 1, 7);
      assertEquals(new LocalDate(2017, 2, 2).toDate(), sheet.getDate(1, 8));
      assertGetDateThrowsException(sheet, 1, 9);
      // row 5
      assertGetDateThrowsException(sheet, 5, 0);
      assertNull(sheet.getDate(5, 1));
      assertNull(sheet.getDate(5, 6));
      assertNull(sheet.getDate(5, 7));
      assertNull(sheet.getDate(5, 8));
      assertNull(sheet.getDate(5, 9));
    }
  }

  static void assertGetDateThrowsException(final Sheet sheet, final int row, final int col) {
    try {
      sheet.getDate(row, col);
      fail("Exception not thrown");
    } catch (final IllegalStateException e) {
      assertEquals("Cannot retrieve a Date from a cell not having type DATE", e.getMessage());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#getBoolean(int, int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetBoolean() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      // row 1
      assertGetBooleanThrowsException(sheet, 1, 0);
      assertGetBooleanThrowsException(sheet, 1, 1);
      assertGetBooleanThrowsException(sheet, 1, 6);
      assertGetBooleanThrowsException(sheet, 1, 7);
      assertGetBooleanThrowsException(sheet, 1, 8);
      assertEquals(true, sheet.getBoolean(1, 9));
      // row 5
      assertGetBooleanThrowsException(sheet, 5, 0);
      assertNull(sheet.getBoolean(5, 1));
      assertNull(sheet.getBoolean(5, 6));
      assertNull(sheet.getBoolean(5, 7));
      assertNull(sheet.getBoolean(5, 8));
      assertNull(sheet.getBoolean(5, 9));
    }
  }

  static void assertGetBooleanThrowsException(final Sheet sheet, final int row, final int col) {
    try {
      sheet.getBoolean(row, col);
      fail("Exception not thrown");
    } catch (final IllegalStateException e) {
      assertEquals("Cannot retrieve a Boolean from a cell not having type BOOLEAN", e.getMessage());
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#getRow(int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetRowInt() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      final Object[] row1 = sheet.getRow(1);
      assertArrayEquals(EXPECTED_ROW_1, row1);
      final Object[] row5 = sheet.getRow(5);
      assertArrayEquals(EXPECTED_ROW_5, row5);
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#getRow(int, int, int)}.
   * 
   * @throws IOException
   * @throws FileNotFoundException
   */
  @Test
  public void testGetRowIntIntInt() throws FileNotFoundException, IOException {
    try (final OdsSpreadsheet spreadsheet = OdsSpreadsheet.open(JFiles.fromResource(RATINGS_ODS))) {
      final Sheet sheet = spreadsheet.getSheet(HEROES);
      final Object[] row1 = sheet.getRow(1, 1, 5);
      for (int i = 0; i < row1.length; i++) {
        assertEquals(EXPECTED_ROW_1[i + 1], row1[i]);
      }
      final Object[] row5 = sheet.getRow(5, 1, 5);
      for (int i = 0; i < row5.length; i++) {
        assertEquals(EXPECTED_ROW_5[i + 1], row5[i]);
      }
    }
  }

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.AbstractSheet#setAutoSize()}.
   */
  @Test
  public void testSetAutoSize() {}

  /**
   * Test method for {@link com.github.mforoni.jspreadsheet.OdsSheet#setAutoSize(int, int)}.
   */
  @Test
  public void testSetAutoSizeIntInt() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setObject(int, int, java.lang.Object)}.
   */
  @Test
  public void testSetObjectIntIntObject() {}

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.OdsSheet#setObject(int, int, java.lang.Object, com.github.mforoni.jspreadsheet.SSCellFormat)}.
   */
  @Test
  public void testSetObjectIntIntObjectSSCellFormat() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setString(int, int, java.lang.String)}.
   */
  @Test
  public void testSetStringIntIntString() {}

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.OdsSheet#setString(int, int, java.lang.String, com.github.mforoni.jspreadsheet.SSCellFormat)}.
   */
  @Test
  public void testSetStringIntIntStringSSCellFormat() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setDouble(int, int, java.lang.Double)}.
   */
  @Test
  public void testSetDoubleIntIntDouble() {}

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.OdsSheet#setDouble(int, int, java.lang.Double, com.github.mforoni.jspreadsheet.SSCellFormat)}.
   */
  @Test
  public void testSetDoubleIntIntDoubleSSCellFormat() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setDate(int, int, java.util.Date)}.
   */
  @Test
  public void testSetDateIntIntDate() {}

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.OdsSheet#setDate(int, int, java.util.Date, com.github.mforoni.jspreadsheet.SSCellFormat)}.
   */
  @Test
  public void testSetDateIntIntDateSSCellFormat() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setBoolean(int, int, java.lang.Boolean)}.
   */
  @Test
  public void testSetBooleanIntIntBoolean() {}

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.OdsSheet#setBoolean(int, int, java.lang.Boolean, com.github.mforoni.jspreadsheet.SSCellFormat)}.
   */
  @Test
  public void testSetBooleanIntIntBooleanSSCellFormat() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#add(int, int, java.lang.Double)}.
   */
  @Test
  public void testAdd() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, java.lang.Object[])}.
   */
  @Test
  public void testSetRowIntObjectArray() {}

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, java.lang.Object[], com.github.mforoni.jspreadsheet.SSCellFormat)}.
   */
  @Test
  public void testSetRowIntObjectArraySSCellFormat() {
    // TODO
  }

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, int, java.lang.Object[])}.
   */
  @Test
  public void testSetRowIntIntObjectArray() {}

  /**
   * Test method for
   * {@link com.github.mforoni.jspreadsheet.AbstractSheet#setRow(int, int, java.lang.Object[], com.github.mforoni.jspreadsheet.SSCellFormat)}.
   */
  @Test
  public void testSetRowIntIntObjectArraySSCellFormat() {
    // TODO
  }
}
