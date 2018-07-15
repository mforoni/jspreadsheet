package com.github.mforoni.jspreadsheet;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import javax.annotation.Nonnull;
import com.github.mforoni.jspreadsheet.Spreadsheets.Mode;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * Wrapping JXL (JEXCEL) API in order to handling Excel files.<br>
 * In general, a wrapper class is any class which "wraps" or "encapsulates" the functionality of
 * another class or component. These are useful by providing a level of abstraction from the
 * implementation of the underlying class or component.
 * <p>
 * To instantiate this class use the methods provided in {@link Spreadsheets} class.
 * <p>
 * <b>Note:</b> JXL API support only <i>xls</i> file, not <i>xlsx</i> (Excel 2007).
 * 
 * @author Foroni Marco
 * @see Spreadsheets
 * @see Spreadsheet
 * @see Sheet
 */
final class JxlSpreadsheet implements Spreadsheet {
  private final File file;
  private final Mode mode;
  private final WritableWorkbook writableWorkbook;
  private final Workbook workbook;

  private JxlSpreadsheet(final File file, final Mode mode) throws IOException {
    this.file = file;
    this.mode = mode;
    switch (mode) {
      case CREATE:
        writableWorkbook = Workbook.createWorkbook(file);
        workbook = null;
        break;
      case OPEN:
        writableWorkbook = null;
        try {
          final WorkbookSettings settings = new WorkbookSettings();
          settings.setRationalization(false);
          settings.setEncoding("Cp1252");
          workbook = Workbook.getWorkbook(file, settings);
        } catch (final BiffException ex) {
          throw new IOException(ex);
        }
        break;
      case EDIT:
        workbook = null;
        Workbook workbook = null;
        try {
          workbook = Workbook.getWorkbook(file);
          writableWorkbook = Workbook.createWorkbook(file, workbook);
        } catch (final BiffException ex) {
          throw new IOException(ex);
        } finally {
          if (workbook != null) {
            workbook.close();
          }
        }
        break;
      default:
        throw new AssertionError();
    }
  }

  static JxlSpreadsheet create(final File file) throws IOException {
    return new JxlSpreadsheet(file, Mode.CREATE);
  }

  /**
   * Opens the specified Excel file in read-only mode.
   * 
   * @param file
   * @throws IOException
   * @see Spreadsheet
   */
  static JxlSpreadsheet open(final File file) throws IOException {
    return new JxlSpreadsheet(file, Mode.OPEN);
  }

  /**
   * Edits the specified Excel file.
   * 
   * @param file
   * @throws IOException
   * @see Spreadsheet
   */
  static JxlSpreadsheet edit(final File file) throws IOException {
    return new JxlSpreadsheet(file, Mode.EDIT);
  }

  /**
   * Checks if the spreadsheet is open in read-only mode, i.e. checks if writableWorkbook is null.
   * 
   * @throws IllegalStateException
   */
  private void canWrite() throws IllegalStateException {
    if (mode == Mode.OPEN || writableWorkbook == null) {
      throw new IllegalStateException(
          "Write operation not allowed on file opened in read-only mode.");
    }
  }

  @Override
  public File getFile() {
    return file;
  }

  /**
   * Returns the sheet from the {@link WritableWorkbook} if write is allowed, otherwise returns the
   * sheet from the {@link Workbook}.
   * 
   * @param sheetIndex
   * @return
   */
  private jxl.Sheet _getSheet(final int sheetIndex) {
    return writableWorkbook != null ? writableWorkbook.getSheet(sheetIndex)
        : workbook.getSheet(sheetIndex);
  }

  /**
   * @param index
   * @return
   * @throws IndexOutOfBoundsException
   * @see Workbook#getSheet(int)
   * @see WritableWorkbook#getSheet(int)
   */
  @Nonnull
  private jxl.Sheet getJxlSheet(final int index) throws IndexOutOfBoundsException {
    if (workbook != null) {
      return workbook.getSheet(index);
    } else if (writableWorkbook != null) {
      return writableWorkbook.getSheet(index);
    } else {
      throw new IllegalStateException();
    }
  }

  /**
   * @param name
   * @return
   * @see Workbook#getSheet(String)
   * @see WritableWorkbook#getSheet(String)
   */
  private jxl.Sheet getJxlSheet(final String name) {
    if (workbook != null) {
      return workbook.getSheet(name);
    } else if (writableWorkbook != null) {
      return writableWorkbook.getSheet(name);
    } else {
      throw new IllegalStateException();
    }
  }

  private String[] _getSheetNames() {
    return writableWorkbook != null ? writableWorkbook.getSheetNames() : workbook.getSheetNames();
  }

  @Override
  public List<String> getSheetNames() {
    return Arrays.asList(_getSheetNames());
  }

  @Nonnull
  @Override
  public List<Sheet> getSheets() {
    final List<Sheet> sheets = new ArrayList<>();
    for (final String name : getSheetNames()) {
      sheets.add(new JxlSheet(getJxlSheet(name), mode.editable()));
    }
    return sheets;
  }

  @Override
  public Sheet getSheet(final String name) {
    return new JxlSheet(getJxlSheet(name), mode.editable());
  }

  @Nonnull
  @Override
  public Sheet getSheet(final int index) {
    return new JxlSheet(getJxlSheet(index), mode.editable());
  }

  @Override
  public Sheet addSheet(final String name) {
    canWrite();
    final int sheets = writableWorkbook.getNumberOfSheets();
    final jxl.Sheet sheet = writableWorkbook.createSheet(name, sheets);
    return new JxlSheet(sheet, mode.editable());
  }

  @Override
  public void write() throws IOException {
    canWrite();
    writableWorkbook.write();
  }

  @Override
  public void close() throws IOException {
    try {
      if (writableWorkbook != null) {
        writableWorkbook.close();
      }
    } catch (final WriteException ex) {
      throw new IOException("Error while closing WritableWorkbook", ex);
    } finally {
      if (workbook != null) {
        workbook.close();
      }
    }
  }
}
