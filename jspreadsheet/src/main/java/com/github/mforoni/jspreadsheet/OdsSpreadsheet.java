package com.github.mforoni.jspreadsheet;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import javax.annotation.Nonnull;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import com.github.mforoni.jspreadsheet.Spreadsheets.Mode;

/**
 * Wrapping jopendocument API in order to handle ODS spreadsheet. To instantiate this class use the
 * methods provided in the {@link Spreadsheets} class.
 * 
 * @author Foroni Marco
 * @see Spreadsheets
 * @see SpreadSheet
 * @see OdsSheet
 */
final class OdsSpreadsheet implements Spreadsheet {
  private final SpreadSheet spreadSheet;
  private final Mode mode;
  private final File file;

  private OdsSpreadsheet(final File file, final Mode mode) throws IOException {
    this.file = file;
    this.mode = mode;
    switch (mode) {
      case CREATE:
        spreadSheet = SpreadSheet.create(1, 1, 1);
        break;
      case EDIT:
        spreadSheet = SpreadSheet.createFromFile(file);
        break;
      case OPEN:
        spreadSheet = SpreadSheet.createFromFile(file);
        break;
      default:
        throw new AssertionError();
    }
  }

  /**
   * Creates a new ODS spreadsheet from scratch placed in the specified {@code file}.
   * 
   * @param file
   * @throws IOException
   * @see Spreadsheet
   */
  static OdsSpreadsheet create(final File file) throws IOException {
    return new OdsSpreadsheet(file, Mode.CREATE);
  }

  /**
   * Edits the specified ODS spreadsheet {@code file}.
   * 
   * @param file an existing ODS spreadsheet file
   * @throws IOException
   * @see Spreadsheet
   */
  static OdsSpreadsheet edit(final File file) throws IOException {
    return new OdsSpreadsheet(file, Mode.EDIT);
  }

  /**
   * Opens the specified ODS spreadsheet {@code file} in read-only mode.
   * 
   * @param file an existing ODS spreadsheet file
   * @throws IOException
   * @see Spreadsheet
   */
  static OdsSpreadsheet open(final File file) throws IOException {
    return new OdsSpreadsheet(file, Mode.OPEN);
  }

  @Override
  public File getFile() {
    return file;
  }

  /**
   * {@inheritDoc}
   */
  @Nonnull
  @Override
  public List<String> getSheetNames() {
    final List<String> names = new ArrayList<>();
    for (int i = 0; i < spreadSheet.getSheetCount(); i++) {
      names.add(spreadSheet.getSheet(i).getName());
    }
    return names;
  }

  /**
   * {@inheritDoc}
   */
  @Nonnull
  @Override
  public List<Sheet> getSheets() {
    final List<Sheet> sheets = new ArrayList<>();
    for (int i = 0; i < spreadSheet.getSheetCount(); i++) {
      final org.jopendocument.dom.spreadsheet.Sheet sheet = spreadSheet.getSheet(i);
      sheets.add(new OdsSheet(sheet));
    }
    return sheets;
  }

  /**
   * {@inheritDoc}
   */
  @Nonnull
  @Override
  public Sheet getSheet(final String name) {
    final org.jopendocument.dom.spreadsheet.Sheet sheet = spreadSheet.getSheet(name);
    if (sheet != null) {
      return new OdsSheet(sheet);
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
    try {
      final org.jopendocument.dom.spreadsheet.Sheet sheet = spreadSheet.getSheet(index);
      return new OdsSheet(sheet);
    } catch (final IndexOutOfBoundsException e) {
      throw new IllegalArgumentException(String.format("Sheet index (%d) is out of range (0..%d)",
          index, spreadSheet.getSheetCount() - 1), e);
    }
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Sheet addSheet(final String name) {
    canWrite();
    // what happens if you add a sheet with an already existing name ? FIXME
    final org.jopendocument.dom.spreadsheet.Sheet sheet = spreadSheet.addSheet(name);
    return new OdsSheet(sheet);
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public void write() throws IOException {
    canWrite();
    spreadSheet.saveAs(file);
  }

  /**
   * {@inheritDoc}
   */
  @Deprecated
  @Override
  public void close() {
    // do nothing
  }

  private void canWrite() throws IllegalStateException {
    if (mode == Mode.OPEN) {
      throw new IllegalStateException(
          "Write operation not allowed on file opened in read-only mode.");
    }
  }
}
