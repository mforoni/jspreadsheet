package com.github.mforoni.jspreadsheet;

import java.util.Date;
import javax.annotation.Nullable;
import org.joda.time.DateTime;

/**
 * @author Foroni Marco
 */
abstract class AbstractSheet implements Sheet {
  @Override
  public int getLastRow() {
    int rowIndex = getRows();
    for (; rowIndex > 0; rowIndex--) {
      final int lastColumn = getLastColumn(rowIndex - 1);
      if (lastColumn != 0) {
        return rowIndex;
      }
    }
    return 0;
  }

  @Override
  public Object[] getRow(final int rowIndex, final int fromColumn, final int toColumn) {
    final int size = toColumn - fromColumn + 1;
    final Object[] values = new Object[size];
    for (int c = 0; c < size; c++) {
      values[c] = getObject(rowIndex, c + fromColumn);
    }
    return values;
  }

  @Override
  public Object[] getRow(final int rowIndex) {
    return getRow(rowIndex, 0, getColumns() - 1);
  }

  @Override
  public void setAutoSize() {
    setAutoSize(0, getColumns());
  }

  @Override
  public void setObject(final int rowIndex, final int columnIndex, @Nullable final Object value) {
    setObject(rowIndex, columnIndex, value, null);
  }

  @Override
  public void setObject(final int rowIndex, final int columnIndex, @Nullable final Object value,
      @Nullable final SSCellFormat cellFormat) {
    if (value != null) {
      if (value instanceof String) {
        setString(rowIndex, columnIndex, (String) value, cellFormat);
      } else if (value instanceof Number) {
        setDouble(rowIndex, columnIndex, ((Number) value).doubleValue(), cellFormat);
      } else if (value instanceof Date) {
        setDate(rowIndex, columnIndex, (Date) value, cellFormat);
      } else if (value instanceof DateTime) {
        final DateTime dateTime = (DateTime) value;
        setDate(rowIndex, columnIndex, dateTime.toDate(), cellFormat);
      } else if (value instanceof Boolean) {
        setBoolean(rowIndex, columnIndex, (Boolean) value, cellFormat);
      } else {
        throw new IllegalStateException(
            String.format("Unable to set value having Type %s: case not handled.",
                value.getClass().getSimpleName()));
      }
    }
  }

  @Override
  public void setString(final int rowIndex, final int columnIndex, final String value) {
    setString(rowIndex, columnIndex, value, null);
  }

  @Override
  public void setDouble(final int rowIndex, final int columnIndex, final Double value) {
    setDouble(rowIndex, columnIndex, value, null);
  }

  @Override
  public void setDate(final int rowIndex, final int columnIndex, final Date value) {
    setDate(rowIndex, columnIndex, value, null);
  }

  @Override
  public void setBoolean(final int rowIndex, final int columnIndex, final Boolean value) {
    setBoolean(rowIndex, columnIndex, value, null);
  }

  @Override
  public void add(final int rowIndex, final int columnIndex, @Nullable final Double value) {
    if (value != null) {
      final Double d = getDouble(rowIndex, columnIndex);
      final double sum = d != null ? d + value : value;
      setDouble(rowIndex, columnIndex, sum);
    }
  }

  @Override
  public void setRow(final int rowIndex, final Object[] values) {
    setRow(rowIndex, 0, values);
  }

  @Override
  public void setRow(final int rowIndex, final Object[] values, final SSCellFormat cellFormat) {
    setRow(rowIndex, 0, values, cellFormat);
  }

  @Override
  public void setRow(final int rowIndex, final int columnOffset, final Object[] values) {
    for (int c = 0; c < values.length; c++) {
      setObject(rowIndex, columnOffset + c, values[c]);
    }
  }

  @Override
  public void setRow(final int rowIndex, final int columnOffset, final Object[] values,
      final SSCellFormat cellFormat) {
    for (int c = 0; c < values.length; c++) {
      setObject(rowIndex, columnOffset + c, values[c], cellFormat);
    }
  }
}
