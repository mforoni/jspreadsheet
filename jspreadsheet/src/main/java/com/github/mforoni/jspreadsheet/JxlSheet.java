package com.github.mforoni.jspreadsheet;

import static jxl.CellType.BOOLEAN;
import static jxl.CellType.BOOLEAN_FORMULA;
import static jxl.CellType.DATE;
import static jxl.CellType.DATE_FORMULA;
import static jxl.CellType.LABEL;
import static jxl.CellType.NUMBER;
import static jxl.CellType.NUMBER_FORMULA;
import static jxl.CellType.STRING_FORMULA;
import java.util.Date;
import javax.annotation.Nullable;
import jxl.BooleanCell;
import jxl.Cell;
import jxl.CellType;
import jxl.CellView;
import jxl.DateCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.format.CellFormat;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

/**
 * @author Foroni Marco
 * @see Sheet
 */
final class JxlSheet extends AbstractSheet {
  private final jxl.Sheet sheet;
  private final boolean editable;

  JxlSheet(final jxl.Sheet sheet, final boolean editable) {
    this.sheet = sheet;
    this.editable = editable;
  }

  @Override
  public String getName() {
    return sheet.getName();
  }

  @Override
  public int getRows() {
    return sheet.getRows();
  }

  @Override
  public int getColumns() {
    return sheet.getColumns();
  }

  @Override
  public int getLastColumn(final int rowIndex) {
    return sheet.getRow(rowIndex).length;
  }

  private WritableSheet getWritableSheet() {
    if (editable && sheet instanceof WritableSheet) {
      return (WritableSheet) sheet;
    } else {
      throw new IllegalStateException(
          "Write operation not allowed on file opened in read-only mode.");
    }
  }

  @Override
  public void setAutoSize(final int fromCol, final int toCol) {
    final WritableSheet sheet = getWritableSheet();
    for (int c = fromCol; c < toCol; c++) {
      final CellView cellView = sheet.getColumnView(c);
      cellView.setAutosize(true);
      sheet.setColumnView(c, cellView);
    }
  }

  public void hideColumn(final int columnIndex) {
    final WritableSheet sheet = getWritableSheet();
    final CellView cellView = sheet.getColumnView(columnIndex);
    cellView.setHidden(true);
    sheet.setColumnView(columnIndex, cellView);
  }

  @Override
  public String getRawValue(final int rowIndex, final int columnIndex) {
    final Cell cell = sheet.getCell(columnIndex, rowIndex);
    return cell.getContents();
  }

  @Override
  public Object getObject(final int rowIndex, final int colIndex) {
    final Cell cell = sheet.getCell(colIndex, rowIndex);
    if (cell.getType().equals(CellType.EMPTY)) {
      return null;
    } else if (cell.getType().equals(NUMBER) || cell.getType().equals(NUMBER_FORMULA)) {
      final NumberCell numberCell = (NumberCell) cell;
      return numberCell.getValue();
    } else if (cell.getType().equals(LABEL) || cell.getType().equals(STRING_FORMULA)) {
      return ((LabelCell) cell).getString();
    } else if (cell.getType().equals(BOOLEAN) || cell.getType().equals(BOOLEAN_FORMULA)) {
      final BooleanCell booleanCell = (BooleanCell) cell;
      return booleanCell.getValue();
    } else if (cell.getType().equals(DATE) || cell.getType().equals(DATE_FORMULA)) {
      final DateCell dateCell = (DateCell) cell;
      return dateCell.getDate();
    } else {
      throw new IllegalStateException("Cell type " + cell.getType() + " not handled.");
    }
  }

  @Override
  public String getString(final int rowIndex, final int colIndex) {
    final Cell cell = sheet.getCell(colIndex, rowIndex);
    if (cell.getType().equals(LABEL) || cell.getType().equals(STRING_FORMULA)) {
      return ((LabelCell) cell).getString();
    } else {
      throw new IllegalStateException(
          String.format("Cannot retrieve a string value from cell [%d, %d] having type %s",
              rowIndex, colIndex, cell.getType()));
    }
  }

  @Override
  public Double getDouble(final int rowIndex, final int colIndex) {
    final Cell cell = sheet.getCell(colIndex, rowIndex);
    if (cell.getType().equals(NUMBER) || cell.getType().equals(NUMBER_FORMULA)) {
      final NumberCell numberCell = (NumberCell) cell;
      return numberCell.getValue();
    } else {
      throw new IllegalStateException(
          String.format("Cannot retrieve a double value from cell [%d, %d] having type %s",
              rowIndex, colIndex, cell.getType()));
    }
  }

  @Override
  public Date getDate(final int rowIndex, final int colIndex) {
    final Cell cell = sheet.getCell(colIndex, rowIndex);
    if (cell.getType().equals(DATE) || cell.getType().equals(DATE_FORMULA)) {
      final DateCell dateCell = (DateCell) cell;
      return dateCell.getDate();
    } else {
      throw new IllegalStateException(
          String.format("Cannot retrieve a double value from cell [%d, %d] having type %s",
              rowIndex, colIndex, cell.getType()));
    }
  }

  @Override
  public Boolean getBoolean(final int rowIndex, final int colIndex) {
    final Cell cell = sheet.getCell(colIndex, rowIndex);
    if (cell.getType().equals(BOOLEAN) || cell.getType().equals(BOOLEAN_FORMULA)) {
      final BooleanCell booleanCell = (BooleanCell) cell;
      return booleanCell.getValue();
    } else {
      throw new IllegalStateException(
          String.format("Cannot retrieve a double value from cell [%d, %d] having type %s",
              rowIndex, colIndex, cell.getType()));
    }
  }

  @Override
  public void setString(final int rowIndex, final int columnIndex, @Nullable final String value,
      @Nullable final SSCellFormat cellFormat) {
    if (value != null) {
      final WritableSheet writableSheet = getWritableSheet();
      WritableCell writableCell;
      try {
        if (cellFormat == null) {
          final Cell cell = writableSheet.getCell(columnIndex, rowIndex);
          final CellFormat format = cell.getCellFormat();
          writableCell = format != null ? new Label(columnIndex, rowIndex, value, format)
              : new Label(columnIndex, rowIndex, value);
        } else {
          writableCell = new Label(columnIndex, rowIndex, value, toWritableCellFormat(cellFormat));
        }
        writableSheet.addCell(writableCell);
      } catch (final WriteException e) {
        throw new IllegalStateException(
            String.format("Error while setting value %s at Sheet=%s, cell=[%d, %d].", value,
                sheet.getName(), rowIndex, columnIndex, e));
      }
    }
  }

  @Override
  public void setDouble(final int rowIndex, final int columnIndex, @Nullable final Double value,
      @Nullable final SSCellFormat cellFormat) {
    if (value != null) {
      final WritableSheet writableSheet = getWritableSheet();
      WritableCell writableCell;
      if (cellFormat == null) {
        final Cell cell = writableSheet.getCell(columnIndex, rowIndex);
        final jxl.format.CellFormat format = cell.getCellFormat();
        writableCell = format != null ? new jxl.write.Number(columnIndex, rowIndex, value, format)
            : new jxl.write.Number(columnIndex, rowIndex, value);
      } else {
        writableCell =
            new jxl.write.Number(columnIndex, rowIndex, value, toWritableCellFormat(cellFormat));
      }
      try {
        writableSheet.addCell(writableCell);
      } catch (final WriteException e) {
        throw new IllegalStateException(
            String.format("Error while setting value %s at Sheet=%s, cell=[%d, %d].", value,
                sheet.getName(), rowIndex, columnIndex, e));
      }
    }
  }

  @Override
  public void setDate(final int rowIndex, final int columnIndex, final Date value,
      final SSCellFormat cellFormat) {
    // TODO Auto-generated method stub
  }

  @Override
  public void setBoolean(final int rowIndex, final int columnIndex, final Boolean value,
      final SSCellFormat cellFormat) {
    // TODO Auto-generated method stub
  }

  private static WritableCellFormat toWritableCellFormat(final SSCellFormat cellFormat) {
    try {
      final WritableCellFormat writableCellFormat =
          new WritableCellFormat(cellFormat.getSSFont().newWritableFont());
      if (cellFormat.getBackgroundColour() != null) {
        writableCellFormat.setBackground(cellFormat.getBackgroundColour().getColour(),
            jxl.format.Pattern.SOLID);
      }
      return writableCellFormat;
    } catch (final WriteException e) {
      throw new IllegalStateException(e);
    }
  }
}
