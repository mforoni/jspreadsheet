package com.github.mforoni.jspreadsheet;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import javax.annotation.Nonnull;
import com.google.common.base.Predicate;
import com.google.common.collect.ImmutableSet;
import com.google.common.collect.Sets;
import com.google.common.io.Files;

/**
 * @author Foroni Marco
 * @see Spreadsheet
 * @see Sheet
 */
public final class Spreadsheets {
  private static final String XLS = "xls";
  private static final String XLSX = "xlsx";
  private static final String ODS = "ods";
  public static final Predicate<Path> IS_SPREADSHEET = new Predicate<Path>() {
    @Override
    public boolean apply(@Nonnull final Path input) {
      final String fileExtension = Files.getFileExtension(input.getFileName().toString());
      return fileExtension.equalsIgnoreCase(XLS) || fileExtension.equalsIgnoreCase(XLSX)
          || fileExtension.equalsIgnoreCase(ODS);
    }
  };

  private Spreadsheets() {
    throw new AssertionError();
  }

  public enum API {
    JXL, POI, SIMPLE_ODF
  }
  enum Mode {
    OPEN, EDIT, CREATE;
    public boolean editable() {
      return this == Mode.OPEN ? false : true;
    }
  }

  /**
   * Returns <tt>true</tt> if the specified file has extension <tt>xls</tt>, <tt>false</tt>
   * otherwise.
   * 
   * @param file
   * @return <tt>true</tt> if the specified file has extension <tt>xls</tt>
   */
  public static boolean isExtensionXls(final File file) {
    return isExtension(file, XLS);
  }

  public static boolean isExtensionXlsx(final File file) {
    return isExtension(file, XLSX);
  }

  public static boolean isExcelFile(final File file) {
    return isExtensionXls(file) || isExtensionXlsx(file);
  }

  private static boolean isExtension(final File file, final String extension) {
    return extension.equalsIgnoreCase(getExtension(file));
  }

  public static String getExtension(final File file) {
    return Files.getFileExtension(file.getName());
  }

  @Nonnull
  public static API detectAPI(final File file) {
    final String extension = getExtension(file);
    if (extension.equalsIgnoreCase(XLS)) {
      return API.POI;
    } else if (extension.equalsIgnoreCase(XLSX)) {
      return API.POI;
    } else if (extension.equalsIgnoreCase(ODS)) {
      return API.SIMPLE_ODF;
    } else {
      throw new IllegalArgumentException(String.format(
          "Unable to detect API for file %s. The extension %s is not valid, please specify a file with one of the following extension: %s, %s, or %s",
          file, extension, XLS, XLSX, ODS));
    }
  }

  @Nonnull
  public static ImmutableSet<API> getAdmittedAPIs(final File file) {
    final String extension = getExtension(file);
    if (extension.equalsIgnoreCase(XLS)) {
      return Sets.immutableEnumSet(API.POI, API.JXL);
    } else if (extension.equalsIgnoreCase(XLSX)) {
      return Sets.immutableEnumSet(API.POI);
    } else if (extension.equalsIgnoreCase(ODS)) {
      return Sets.immutableEnumSet(API.SIMPLE_ODF);
    } else {
      return ImmutableSet.of();
    }
  }

  private static void validateAPI(final File file, final API api) {
    final ImmutableSet<API> admittedAPIs = getAdmittedAPIs(file);
    if (!admittedAPIs.contains(api)) {
      if (admittedAPIs.size() > 0) {
        throw new IllegalStateException(
            String.format("Unable to open the file %s with API %s. Please use this API %s", file,
                api, admittedAPIs));
      } else {
        throw new IllegalStateException(String.format(
            "Unable to open the file %s. Please specify a file with one of the following extension: %s, %s, or %s",
            file, XLS, XLSX, ODS));
      }
    }
  }

  /**
   * Creates a new {@link Spreadsheet} placed in the specified output {@code file}.
   * 
   * @param file the output {@code File}
   * @return an {@link Spreadsheet} opened in write-mode
   * @throws IOException
   * @see Spreadsheet
   */
  public static Spreadsheet create(final File file) throws IOException {
    return _create(file, detectAPI(file));
  }

  /**
   * Creates a {@link Spreadsheet} placed in the given output {@code file} using the specified
   * {@link API}.
   * 
   * @param file the output {@code File}
   * @param api the API to use
   * @return an {@link Spreadsheet} opened in write-mode
   * @throws IOException
   * @see Spreadsheet
   * @see API
   */
  public static Spreadsheet create(final File file, final API api) throws IOException {
    validateAPI(file, api);
    return _create(file, api);
  }

  private static Spreadsheet _create(final File file, final API api) throws IOException {
    switch (api) {
      case JXL:
        return JxlSpreadsheet.create(file);
      case POI:
        return PoiSpreadsheet.create(file);
      case SIMPLE_ODF:
        return OdsSpreadsheet.create(file);
      default:
        // should never occur
        throw new IllegalArgumentException("API " + api + " not handled");
    }
  }

  public static Spreadsheet edit(final File in) throws IOException {
    return _edit(in, detectAPI(in));
  }

  /**
   * Returns a new {@link Spreadsheet} opened in write-mode given an input {@code File}.
   *
   * @param in the input file
   * @param api the API to use
   * @return a {@link Spreadsheet} opened in write-mode
   * @throws IOException
   * @see Spreadsheet
   * @see API
   */
  public static Spreadsheet edit(final File in, final API api) throws IOException {
    validateAPI(in, api);
    return _edit(in, api);
  }

  private static Spreadsheet _edit(final File in, final API api) throws IOException {
    switch (api) {
      case JXL:
        return JxlSpreadsheet.edit(in);
      case POI:
        return PoiSpreadsheet.edit(in);
      case SIMPLE_ODF:
        return OdsSpreadsheet.edit(in);
      default:
        // should never occur
        throw new IllegalArgumentException("API " + api + " not handled");
    }
  }

  /**
   * Returns a {@link Spreadsheet} opened in read-only mode from the given input file {@code in}.
   *
   * @param in the input file
   * @return an {@link Spreadsheet} opened in read-only mode
   * @throws IOException
   * @see Spreadsheet
   */
  public static Spreadsheet open(final File in) throws IOException {
    return _open(in, detectAPI(in));
  }

  /**
   * Returns a {@link Spreadsheet} opened in read-only from the given input file {@code in} and
   * {@link API}.
   *
   * @param in the input file
   * @param api the API to use
   * @return a {@link Spreadsheet} opened in read-only mode
   * @throws IOException
   * @see Spreadsheet
   * @see API
   */
  public static Spreadsheet open(final File in, final API api) throws IOException {
    validateAPI(in, api);
    return _open(in, api);
  }

  private static Spreadsheet _open(final File in, final API api) throws IOException {
    if (!in.exists()) {
      throw new IOException("File " + in.getName() + " not found.");
    }
    switch (api) {
      case JXL:
        return JxlSpreadsheet.open(in);
      case POI:
        return PoiSpreadsheet.open(in);
      case SIMPLE_ODF:
        return OdsSpreadsheet.open(in);
      default:
        // should never occur
        throw new IllegalStateException("Undefined Excel API Class type.");
    }
  }

  public static int getColumnIndex(final String columnLabel) {
    int columnIndex = 0;
    int esp = 0;
    for (int i = columnLabel.length() - 1; i >= 0; i--) {
      final char ch = Character.toUpperCase(columnLabel.charAt(i));
      final int value = ch - 'A' + 1;
      columnIndex += (int) Math.pow(26, esp++) * value;
    }
    return columnIndex;
  }
}
