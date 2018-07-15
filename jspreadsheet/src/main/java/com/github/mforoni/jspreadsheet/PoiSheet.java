package com.github.mforoni.jspreadsheet;

import java.util.Date;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Foroni Marco
 * @see Sheet
 */
final class PoiSheet extends AbstractSheet {
  private final org.apache.poi.ss.usermodel.Sheet sheet;
  private final Workbook workbook;
  private final boolean editable;

  PoiSheet(final org.apache.poi.ss.usermodel.Sheet sheet, final Workbook workbook,
      final boolean editable) {
    this.sheet = sheet;
    this.workbook = workbook;
    this.editable = editable;
  }

  private void canWrite() throws IllegalStateException {
    if (!editable) {
      throw new IllegalStateException(
          "Write operation not allowed on file opened in read-only mode.");
    }
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public String getName() {
    return sheet.getSheetName();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getRows() {
    // return sheet.getPhysicalNumberOfRows();
    return sheet.getLastRowNum() + 1;
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getColumns() {
    int maxNumCols = 0;
    for (int r = 0; r < getRows(); r++) {
      final int lastColumn = getLastColumn(r);
      if (lastColumn > maxNumCols) {
        maxNumCols = lastColumn;
      }
    }
    return maxNumCols;
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getLastColumn(final int rowIndex) {
    final Row row = sheet.getRow(rowIndex);
    if (row != null) {
      return row.getLastCellNum(); // short
    }
    return 0;
  }

  @Override
  public void setAutoSize(final int fromCol, final int toCol) {
    canWrite();
    for (int c = fromCol; c < toCol; c++) {
      sheet.autoSizeColumn(c);
    }
  }

  public void hideColumn(final int columnIndex) {
    canWrite();
    sheet.setColumnHidden(columnIndex, true);
  }

  private boolean isFormula(final int rowIndex, final int columnIndex) {
    final Cell cell = getCell(sheet, rowIndex, columnIndex);
    if (cell != null && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
      return true;
    }
    return false;
  }

  @Nullable
  private static Cell getCell(final org.apache.poi.ss.usermodel.Sheet sheet, final int rowIndex,
      final int columnIndex) {
    final Row row = sheet.getRow(rowIndex);
    return row != null ? row.getCell(columnIndex) : null;
  }

  @Override
  public String getRawValue(final int rowIndex, final int columnIndex) {
    final DataFormatter df = new DataFormatter();
    final Cell cell = getCell(sheet, rowIndex, columnIndex);
    return df.formatCellValue(cell);
  }

  @Override
  public Object getObject(final int rowIndex, final int columnIndex) {
    final Cell cell = getCell(sheet, rowIndex, columnIndex);
    if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
      return null;
    } else {
      final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      final CellValue cellValue = evaluator.evaluate(cell);
      switch (cellValue.getCellType()) {
        case Cell.CELL_TYPE_BOOLEAN:
          return cellValue.getBooleanValue();
        case Cell.CELL_TYPE_NUMERIC:
          if (DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue();
          }
          return cellValue.getNumberValue();
        case Cell.CELL_TYPE_STRING:
          return cellValue.getStringValue();
        case Cell.CELL_TYPE_BLANK:
          return null;
        case Cell.CELL_TYPE_ERROR:
          return cellValue.getErrorValue();
        case Cell.CELL_TYPE_FORMULA:
          // CELL_TYPE_FORMULA should never occur
          throw new AssertionError("Cell having type FORMULA after evaluation not expected");
        default:
          throw new IllegalStateException("Cell type " + cell.getCellType() + " not handled");
      }
    }
  }

  @Nullable
  @Override
  public String getString(final int rowIndex, final int columnIndex) {
    final Cell cell = getCell(sheet, rowIndex, columnIndex);
    return cell != null ? cell.getStringCellValue() : null;
  }

  @Override
  public Double getDouble(final int rowIndex, final int columnIndex) {
    final Cell cell = getCell(sheet, rowIndex, columnIndex);
    return cell != null ? cell.getNumericCellValue() : null;
  }

  @Override
  public Date getDate(final int rowIndex, final int columnIndex) {
    final Cell cell = getCell(sheet, rowIndex, columnIndex);
    return cell != null ? cell.getDateCellValue() : null;
  }

  @Override
  public Boolean getBoolean(final int rowIndex, final int columnIndex) {
    final Cell cell = getCell(sheet, rowIndex, columnIndex);
    return cell != null ? cell.getBooleanCellValue() : null;
  }

  @Nonnull
  private static Cell getOrCreateCell(final org.apache.poi.ss.usermodel.Sheet sheet,
      final int rowIndex, final int columnIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    Cell cell = row.getCell(columnIndex);
    if (cell == null) {
      cell = row.createCell(columnIndex);
    }
    return cell;
  }

  @Override
  public void setString(final int rowIndex, final int columnIndex, @Nullable final String value,
      @Nullable final SSCellFormat cellFormat) {
    canWrite();
    if (value != null) {
      final Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
      // set cell type and the value
      cell.setCellType(Cell.CELL_TYPE_STRING);
      cell.setCellValue(value);
      if (cellFormat != null) {
        setCellFormat(cell, cellFormat);
      }
    }
  }

  @Override
  public void setDouble(final int rowIndex, final int columnIndex, @Nullable final Double value,
      @Nullable final SSCellFormat cellFormat) {
    canWrite();
    if (value != null) {
      final Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
      // set cell type and the value
      cell.setCellType(Cell.CELL_TYPE_NUMERIC);
      cell.setCellValue(value);
      if (cellFormat != null) {
        setCellFormat(cell, cellFormat);
      }
    }
  }

  @Override
  public void setDate(final int rowIndex, final int columnIndex, @Nullable final Date value,
      @Nullable final SSCellFormat cellFormat) {
    canWrite();
    if (value != null) {
      final Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
      // set cell type and the value
      final CellStyle cellStyle = workbook.createCellStyle();
      final CreationHelper createHelper = workbook.getCreationHelper();
      cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
      cell.setCellStyle(cellStyle);
      cell.setCellValue(value);
      if (cellFormat != null) {
        setCellFormat(cell, cellFormat);
      }
    }
  }

  @Override
  public void setBoolean(final int rowIndex, final int columnIndex, @Nullable final Boolean value,
      @Nullable final SSCellFormat cellFormat) {
    canWrite();
    if (value != null) {
      final Cell cell = getOrCreateCell(sheet, rowIndex, columnIndex);
      // set cell type and the value
      cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
      cell.setCellValue(value);
      if (cellFormat != null) {
        setCellFormat(cell, cellFormat);
      }
    }
  }

  private void setCellFormat(@Nonnull final Cell cell, @Nonnull final SSCellFormat cellFormat) {
    final Font font = workbook.createFont();
    final CellStyle cellStyle = workbook.createCellStyle();
    final SSFont ssFont = cellFormat.getSSFont();
    font.setFontName(ssFont.getName());
    font.setFontHeightInPoints((short) ssFont.getSize());
    if (ssFont.isBold()) {
      // font.setBold(true);
      // or:
      font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    }
    cellStyle.setFont(font);
    final SSColor backgroundColour = cellFormat.getBackgroundColour();
    if (backgroundColour != null) {
      cellStyle.setFillForegroundColor(backgroundColour.getPoiIndex());
      cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
    }
    cell.setCellStyle(cellStyle);
  }
}
