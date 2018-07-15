package com.github.mforoni.jspreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import javax.annotation.Nonnull;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.github.mforoni.jspreadsheet.Spreadsheets.Mode;

/**
 * Wrapping APACHE POI API in order to handle Microsoft Excel Files. To instantiate this class use
 * the methods provided in the {@link Spreadsheets} class.
 * <p>
 * <b>Note</b>: the APACHE POI class supports both <i>xls</i> and <i>xlsx</i> Excel files.
 * 
 * @see Spreadsheets
 * @see Spreadsheet
 * @see PoiSheet
 * @author Foroni Marco
 */
final class PoiSpreadsheet implements Spreadsheet {
  private final Workbook workbook;
  private final File file;
  private final Mode mode;

  private PoiSpreadsheet(final File file, final Mode mode) throws IOException {
    if (file == null) {
      throw new IllegalArgumentException("A file must be specified");
    }
    this.file = file;
    this.mode = mode;
    switch (mode) {
      case OPEN:
        try {
          workbook = WorkbookFactory.create(file, null, true);
        } catch (final EncryptedDocumentException | InvalidFormatException ex) {
          throw new IOException(ex);
        }
        break;
      case EDIT:
        try (final FileInputStream fis = new FileInputStream(file)) {
          workbook = WorkbookFactory.create(fis);
        } catch (final EncryptedDocumentException | InvalidFormatException ex) {
          throw new IOException(ex);
        }
        break;
      case CREATE:
        if (file.exists()) {
          throw new IllegalStateException(String.format(
              "Cannot create file %s: the file already exist", file));
        }
        com.google.common.io.Files.createParentDirs(file);
        workbook = Spreadsheets.isExtensionXlsx(file) ? new XSSFWorkbook() : new HSSFWorkbook();
        break;
      default:
        throw new AssertionError();
    }
  }

  /**
   * Creates a new Excel file from scratch placed in the specified {@code file}.
   * 
   * @param file
   * @throws IOException
   * @see Spreadsheet
   */
  static PoiSpreadsheet create(final File file) throws IOException {
    return new PoiSpreadsheet(file, Mode.CREATE);
  }

  /**
   * Edits the specified Excel {@code file}.
   * 
   * @param file
   * @throws IOException
   * @see Spreadsheet
   */
  static PoiSpreadsheet edit(final File file) throws IOException {
    return new PoiSpreadsheet(file, Mode.EDIT);
  }

  /**
   * Opens the specified Excel {@code file} in read-only mode.
   * 
   * @param file
   * @throws IOException
   * @see Spreadsheet
   */
  static PoiSpreadsheet open(final File file) throws IOException {
    return new PoiSpreadsheet(file, Mode.OPEN);
  }

  @Override
  public File getFile() {
    return file;
  }

  @Override
  public List<String> getSheetNames() {
    final List<String> names = new ArrayList<>();
    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      names.add(workbook.getSheetName(i));
    }
    return names;
  }

  @Override
  public List<Sheet> getSheets() {
    final List<Sheet> sheets = new ArrayList<>();
    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      sheets.add(new PoiSheet(workbook.getSheetAt(i), workbook, mode.editable()));
    }
    return sheets;
  }

  /**
   * {@inheritDoc}
   */
  @Nonnull
  @Override
  public Sheet getSheet(final String name) {
    final org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(name);
    if (sheet != null) {
      return new PoiSheet(sheet, workbook, mode.editable());
    } else {
      throw new IllegalArgumentException("Cannot find sheet having name " + name);
    }
  }

  /**
   * {@inheritDoc}
   */
  @Nonnull
  @Override
  public Sheet getSheet(final int index) {
    return new PoiSheet(workbook.getSheetAt(index), workbook, mode.editable());
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Sheet addSheet(final String name) {
    canWrite();
    final org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet(name);
    return new PoiSheet(sheet, workbook, mode.editable());
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public void write() throws IOException {
    canWrite();
    try (final FileOutputStream fos = new FileOutputStream(file)) {
      workbook.write(fos);
    }
  }

  private void canWrite() throws IllegalStateException {
    if (mode == Mode.OPEN) {
      throw new IllegalStateException(
          "Write operation not allowed on file opened in read-only mode.");
    }
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public void close() throws IOException {
    if (workbook != null) {
      workbook.close();
    }
  }
}
