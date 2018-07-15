package com.github.mforoni.jspreadsheet;

import java.io.File;
import java.io.IOException;
import java.util.List;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;

/**
 * Represents a generic Spreadsheet: i.e. an <i>ODF Spreadsheet</i> or a <i>Microsoft Excel</i>
 * file.
 * <p>
 * The Open Document Format for Office Applications (ODF), also known as OpenDocument, is an
 * XML-based file format for spreadsheets, charts, presentations and word processing documents. It
 * was developed with the aim of providing an open, XML-based file format specification for office
 * applications.
 * <p>
 * <b>Microsoft Excel Legacy</b><br>
 * Legacy filename extensions denote binary Microsoft Excel formats that are deprecated with the
 * release of Microsoft Office 2007. Although the latest version of Microsoft Excel can still open
 * them, they are no longer developed. Legacy filename extensions include:
 * <ul>
 * <li><i>xls</i> - Legacy Excel worksheets; officially designated "Microsoft Excel 97-2003
 * Worksheet"</li>
 * <li><i>xlt</i> - Legacy Excel templates; officially designated "Microsoft Excel 97-2003
 * Template"</li>
 * <li><i>xlm</i> - Excel macro</li>
 * </ul>
 * <b>OOXML</b><br>
 * Office Open XML (OOXML) format was introduced with Microsoft Office 2007 and became the default
 * format of Microsoft Excel ever since. Excel-related file extensions of this format include:
 * <ul>
 * <li><i>xlsx</i> - Excel workbook</li>
 * <li><i>xlsm</i> - Excel macro-enabled workbook; same as xlsx but may contain macros and
 * scripts</li>
 * <li><i>xltx</i> - Excel template</li>
 * <li><i>xltm</i> - Excel macro-enabled template; same as xltx but may contain macros and
 * scripts</li>
 * </ul>
 * <b>Other formats</b><br>
 * Microsoft Excel uses dedicated file format that are not part of OOXML and use the following
 * extensions:
 * <ul>
 * <li>xlsb - Excel binary worksheet (BIFF12)</li>
 * <li>xla - Excel add-in or macro</li>
 * <li>xlam - Excel add-in</li>
 * <li>xll - Excel XLL add-in; a form of DLL-based add-in[1]</li>
 * <li>xlw - Excel workspace; previously known as "workbook"</li>
 * </ul>
 * <p>
 * To instantiate this class use the methods provided in {@link Office} class.
 * <p>
 * Limits:
 * <ul>
 * <li>Version: (a) = Max. Rows (b) = Max. Columns (c) = Max. Cols by letter
 * <li>Excel 365*: (a) = 1,048,576 (b) = 16,384 (c) = XFD
 * <li>Excel 2013: (a) = 1,048,576 (b) = 16,384 (c) = XFD
 * <li>Excel 2010: (a) = 1,048,576 (b) = 16,384 (c) = XFD
 * <li>Excel 2007: (a) = 1,048,576 (b) = 16,384 (c) = XFD
 * <li>Excel 2003: (a) = 65,536 (b) = 256 (c) = IV
 * <li>Excel 2002 (XP): (a) = 65,536 (b) = 256 (c) = IV
 * <li>Excel 2000: (a) = 65,536 (b) = 256 (c) = IV
 * <li>Excel 97: (a) = 65,536 (b) = 256 (c) = IV
 * <li>Excel 95: (a) = 16,384 (b) = 256 (c) = IV
 * <li>Excel 5: (a) = 16,384 (b) = 256 (c) = IV
 * </ul>
 * see <a href=
 * "http://superuser.com/questions/366468/what-is-the-maximum-allowed-rows-in-a-microsoft-excel-xls-or-xlsx">Excel
 * rows and columns limits</a>
 * 
 * @author Foroni Marco
 */
public interface Spreadsheet extends AutoCloseable {
  public static final int EXCEL_2013_ROW_LIMIT = 1048576;
  public static final int EXCEL_2013_COLUMN_LIMIT = 16384;
  public static final int EXCEL_2010_ROW_LIMIT = 1048576;
  public static final int EXCEL_2010_COLUMN_LIMIT = 16384;
  public static final int EXCEL_2007_ROW_LIMIT = 1048576;
  public static final int EXCEL_2007_COLUMN_LIMIT = 16384;
  public static final int EXCEL_2003_ROW_LIMIT = 65536;
  public static final int EXCEL_2003_COLUMN_LIMIT = 256;

  @Nonnull
  public File getFile();

  /**
   * Returns the sheet names contained in the spreadsheet.
   * 
   * @return all the sheet names
   */
  @Nonnull
  public List<String> getSheetNames();

  /**
   * Returns the sheets contained in the spreadsheet.
   * 
   * @return all the sheet names
   * @see Sheet
   */
  @Nonnull
  public List<Sheet> getSheets();

  /**
   * Returns the {@link Sheet} given the sheet name. Throws an {@link IllegalArgumentException} if
   * there's no sheet having name equals to the specified name.
   * 
   * @param name the sheet name
   * @return the {@link Sheet} having name equals to the specified {@code name}
   * @throws IllegalArgumentException if there's no sheet having the specified name
   * @see Sheet
   */
  @Nonnull
  public Sheet getSheet(final String name) throws IllegalArgumentException;

  /**
   * Returns the {@link Sheet} given the sheet index.Throws an {@link IllegalArgumentException} if
   * there's no sheet having the specified index.
   * 
   * @param index the sheet index
   * @return the {@link Sheet} having the specified {@code index}
   * @throws IllegalArgumentException if there's no sheet having the specified index
   * @see Sheet
   */
  @Nullable
  public Sheet getSheet(final int index) throws IllegalArgumentException;

  /**
   * Add a new sheet having the specified name. Throws an {@link IllegalStateException} if already
   * exist a sheet having the specified name.
   * 
   * @param name the sheet name
   * @return the newly created {@link Sheet}
   * @throws IllegalStateException if already exist a sheet having the specified name
   * @see Sheet
   */
  @Nonnull
  public Sheet addSheet(final String name);

  public void write() throws IOException;

  @Override
  public void close() throws IOException;
}
