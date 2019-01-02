package com.github.mforoni.jspreadsheet;

import java.math.BigDecimal;
import java.util.Date;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.xml.datatype.Duration;
import org.jopendocument.dom.ODValueType;
import org.jopendocument.dom.spreadsheet.Cell;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import com.google.common.base.Preconditions;

final class OdsSheet extends AbstractSheet {
  private final Sheet sheet;

  OdsSheet(final Sheet sheet) {
    this.sheet = sheet;
  }

  private Cell<SpreadSheet> getImmutableCell(final int rowIndex, final int columnIndex) {
    return sheet.getImmutableCellAt(columnIndex, rowIndex);
  }

  private MutableCell<SpreadSheet> getMutableCell(final int rowIndex, final int columnIndex) {
    return sheet.getCellAt(columnIndex, rowIndex);
  }

  private ODValueType getCellType(final int rowIndex, final int columnIndex) {
    return sheet.getImmutableCellAt(columnIndex, rowIndex).getValueType();
  }

  /**
   * {@inheritDoc}
   */
  @Nonnull
  @Override
  public String getName() {
    return sheet.getName();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getRows() {
    return sheet.getRowCount();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getColumns() {
    return sheet.getColumnCount();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getLastColumn(final int rowIndex) {
    int lastCol = sheet.getColumnCount();
    for (; lastCol > 0; lastCol--) {
      if (getObject(rowIndex, lastCol - 1) != null) {
        return lastCol;
      }
    }
    return 0;
  }

  /**
   * {@inheritDoc}
   */
  @Nonnull
  @Override
  public String getRawValue(final int rowIndex, final int columnIndex) {
    final Cell<SpreadSheet> immutableCell = getImmutableCell(rowIndex, columnIndex);
    return immutableCell.getFormula() != null ? immutableCell.getFormula()
        : immutableCell.getTextValue();
  }

  /**
   * {@inheritDoc}
   */
  @Nullable
  @Override
  public Object getObject(final int rowIndex, final int columnIndex) {
    if (getCellType(rowIndex, columnIndex) == null) {
      return null;
    }
    return getImmutableCell(rowIndex, columnIndex).getValue();
  }

  @Nullable
  @Override
  public String getString(final int rowIndex, final int columnIndex) {
    final ODValueType cellType = getCellType(rowIndex, columnIndex);
    if (cellType == null) {
      return null;
    } else if (cellType == ODValueType.STRING) {
      return (String) getImmutableCell(rowIndex, columnIndex).getValue();
    } else {
      throw new IllegalStateException(
          "Cannot retrieve a String from a cell not having type STRING");
    }
  }

  @Override
  public Double getDouble(final int rowIndex, final int columnIndex)
      throws IllegalStateException, NumberFormatException {
    final ODValueType cellType = getCellType(rowIndex, columnIndex);
    if (cellType == null) {
      return null;
    } else if (cellType == ODValueType.FLOAT || cellType == ODValueType.CURRENCY
        || cellType == ODValueType.PERCENTAGE) {
      final Object value = getImmutableCell(rowIndex, columnIndex).getValue();
      Preconditions.checkArgument(value instanceof BigDecimal);
      final BigDecimal bd = (BigDecimal) value;
      return bd.doubleValue();
    } else {
      throw new IllegalStateException(
          "Cannot retrieve a Double from a cell not having type FLOAT, CURRENCY or PERCENTAGE");
    }
  }

  @Override
  public Date getDate(final int rowIndex, final int columnIndex) {
    final ODValueType cellType = getCellType(rowIndex, columnIndex);
    if (cellType == null) {
      return null;
    } else if (cellType == ODValueType.DATE) {
      final Object value = getImmutableCell(rowIndex, columnIndex).getValue();
      Preconditions.checkArgument(value instanceof Date);
      return (Date) value;
    } else {
      throw new IllegalStateException("Cannot retrieve a Date from a cell not having type DATE");
    }
  }

  /**
   * FIXME
   * 
   * @param rowIndex
   * @param columnIndex
   * @return
   */
  public Duration getDuration(final int rowIndex, final int columnIndex) {
    final ODValueType cellType = getCellType(rowIndex, columnIndex);
    if (cellType == null) {
      return null;
    } else if (cellType == ODValueType.TIME) {
      final Object value = getImmutableCell(rowIndex, columnIndex).getValue();
      Preconditions.checkArgument(value instanceof Duration);
      final Duration duration = (Duration) value;
      return duration;
    } else {
      throw new IllegalStateException(
          "Cannot retrieve a Duration from a cell not having type DURATION");
    }
  }

  @Override
  public Boolean getBoolean(final int rowIndex, final int columnIndex) {
    final ODValueType cellType = getCellType(rowIndex, columnIndex);
    if (cellType == null) {
      return null;
    } else if (cellType == ODValueType.BOOLEAN) {
      return (Boolean) getImmutableCell(rowIndex, columnIndex).getValue();
    } else {
      throw new IllegalStateException(
          "Cannot retrieve a Boolean from a cell not having type BOOLEAN");
    }
  }

  @Override
  public void setAutoSize(final int fromColumn, final int toColumn) {
    throw new IllegalStateException("Feature not avaiable");
  }

  @Override
  public void setObject(final int rowIndex, final int columnIndex, @Nullable final Object value,
      @Nullable final SSCellFormat cellFormat) {
    if (value != null) {
      if (columnIndex + 1 > getColumns()) {
        sheet.setColumnCount(columnIndex + 1);
      }
      if (rowIndex + 1 > getRows()) {
        sheet.setRowCount(rowIndex + 1);
      }
      final MutableCell<SpreadSheet> cell = getMutableCell(rowIndex, columnIndex);
      cell.setValue(value);
      if (cellFormat != null) {
        setCellStyle(cell, cellFormat);
      }
    }
  }

  private static void setCellStyle(final MutableCell<SpreadSheet> cell,
      final SSCellFormat cellFormat) {
    throw new IllegalStateException("Feature not avaiable"); // TODO
  }

  @Override
  public void setString(final int rowIndex, final int columnIndex, @Nullable final String value,
      @Nullable final SSCellFormat cellFormat) {
    setObject(rowIndex, columnIndex, value, cellFormat);
  }

  @Override
  public void setDouble(final int rowIndex, final int columnIndex, @Nullable final Double value,
      @Nullable final SSCellFormat cellFormat) {
    setObject(rowIndex, columnIndex, value, cellFormat);
  }

  @Override
  public void setDate(final int rowIndex, final int columnIndex, @Nullable final Date value,
      @Nullable final SSCellFormat cellFormat) {
    setObject(rowIndex, columnIndex, value, cellFormat);
  }

  @Override
  public void setBoolean(final int rowIndex, final int columnIndex, @Nullable final Boolean value,
      @Nullable final SSCellFormat cellFormat) {
    setObject(rowIndex, columnIndex, value);
  }
}
