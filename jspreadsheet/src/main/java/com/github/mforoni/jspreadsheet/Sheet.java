package com.github.mforoni.jspreadsheet;

import java.util.Date;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import com.github.mforoni.jspreadsheet.Spreadsheets.API;
import com.google.common.annotations.Beta;

/**
 * @author Foroni Marco
 */
public interface Sheet {
  @Nonnull
  public String getName();

  /**
   * Returns the number of rows in the sheet (according to the underlying {@link API}).
   * 
   * @return the number of rows
   */
  public int getRows();

  @Beta
  public int getLastRow();

  /**
   * Returns the number of columns in the sheet (according to the underlying {@link API}).
   * 
   * @return the number of columns
   */
  public int getColumns();

  /**
   * Returns the last written column at the row having the specified index. Returns 0 if no column
   * is written.
   * 
   * @param rowIndex index of the row (starting from 0)
   * @return the last written column
   */
  @Beta
  public int getLastColumn(final int rowIndex);

  @Beta
  public enum ValueType {
  }

  // @Beta
  // public ValueType getValueType(final int rowIndex, final int columnIndex); // FIXME
  /**
   * Returns the string content of the specified cell. An empty cell will be returned as an empty
   * string. Formulas in formula type cells will not be evaluated.
   * 
   * @param rowIndex
   * @param columnIndex
   * @return the string content of the specified cell
   */
  @Nonnull
  @Beta
  public String getRawValue(final int rowIndex, final int columnIndex);

  /**
   * Retrieves the value of the cell at [{@code rowIndex}, {@code columnIndex}]. The returned value
   * type depends on the underlying API. // FIXME improve
   * 
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @return the value of the cell
   */
  @Nullable
  public Object getObject(final int rowIndex, int columnIndex);

  /**
   * Retrieves the {@code String} value of the cell at [{@code rowIndex}, {@code columnIndex}]. For
   * blank cells it returns <tt>null</tt>. For numeric cells and formula Cells that are not string
   * Formulas it throws an exception. Formulas in formula type cells will be evaluated.
   * 
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @return the string value of the cell at the specified position
   */
  @Nullable
  public String getString(final int rowIndex, final int columnIndex);

  @Nullable
  public Date getDate(final int rowIndex, int columnIndex);

  @Nullable
  public Boolean getBoolean(final int rowIndex, int columnIndex);

  /**
   * Retrieves the {@code Double} value of the cell at [{@code rowIndex}, {@code columnIndex}]. For
   * blank cells it returns <tt>null</tt>. For text type cell it throws an exception. Formulas in
   * formula type cells will be evaluated.
   *
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @return the value of the cell as a double type
   * @throws IllegalStateException if the cell type is not numeric
   * @throws NumberFormatException if the cell value isn't a parsable double.
   */
  @Nullable
  public Double getDouble(final int rowIndex, final int columnIndex)
      throws IllegalStateException, NumberFormatException;

  @Nonnull
  public Object[] getRow(final int rowIndex);

  @Nonnull
  public Object[] getRow(final int rowIndex, final int fromColumn, final int toColumn);

  @Beta
  public void setAutoSize(final int fromColumn, final int toColumn);

  @Beta
  public void setAutoSize();

  public void setObject(final int rowIndex, final int columnIndex, @Nullable final Object value);

  public void setObject(final int rowIndex, final int columnIndex, @Nullable final Object value,
      @Nullable final SSCellFormat cellFormat);

  /**
   * Set the cell at [{@code rowIndex}, {@code columnIndex}] with the specified string.
   * 
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @param value a {@code String} that contains the value to be set in the cell
   */
  public void setString(final int rowIndex, final int columnIndex, @Nullable final String value);

  /**
   * Set the cell at [{@code rowIndex}, {@code columnIndex}] with the specified string and cell
   * format. If {@code cellFormat} is <tt>null</tt> keeps the current cell format.
   * 
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @param value the {@code string} value to be set in the cell
   * @param cellFormat the format of the cell
   * @see SSCellFormat
   */
  public void setString(final int rowIndex, final int columnIndex, @Nullable final String value,
      @Nullable final SSCellFormat cellFormat);

  /**
   * Set the cell at [{@code rowIndex}, {@code columnsIndex}] with the specified double
   * {@code value} keeping the current cell format.
   * 
   * @param sheetIndex index of the sheet (starting from 0)
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @param value the {@code Double} value to be set in the cell
   */
  public void setDouble(final int rowIndex, final int columnIndex, @Nullable final Double value);

  /**
   * Set the cell at [{@code rowIndex}, {@code columnIndex}] with the specified double {@code value}
   * and the given cell format. If {@code cellFormat} is <tt>null</tt> keeps the current cell
   * format.
   *
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @param value the {@code Double} value to be set in the cell
   * @param cellFormat the format of the cell
   */
  public void setDouble(final int rowIndex, final int columnIndex, @Nullable final Double value,
      @Nullable final SSCellFormat cellFormat);

  @Beta
  public void setDate(final int rowIndex, final int columnIndex, @Nullable final Date value); // or
                                                                                              // DateTime

  @Beta
  public void setDate(final int rowIndex, final int columnIndex, @Nullable final Date value,
      @Nullable final SSCellFormat cellFormat);

  public void setBoolean(final int rowIndex, final int columnIndex, @Nullable final Boolean value);

  public void setBoolean(final int rowIndex, final int columnIndex, @Nullable final Boolean value,
      @Nullable final SSCellFormat cellFormat);

  /**
   * Add a double {@code value} to the current value of the cell at[{@code rowIndex},
   * {@code columnIndex}] keeping the current cell format.
   * 
   * @param rowIndex of the cell (starting from 0)
   * @param columnIndex of the cell (starting from 0)
   * @param value the {@code Double} value to be add in the cell
   */
  @Beta
  public void add(final int rowIndex, final int columnIndex, @Nullable final Double value);

  /**
   * Set the row having the specified index with the provided values keeping the current cell
   * format.
   *
   * @param rowIndex index of the row (starting from 0)
   * @param values the values that has to be set
   */
  public void setRow(final int rowIndex, final Object[] values);

  /**
   * Set the row having the specified index with the provided values starting from the column
   * {@code columnOffset} and keeping the current cell format.
   * 
   * @param rowIndex of the row
   * @param columnOffset
   * @param values the values that has to be set
   */
  public void setRow(final int rowIndex, final int columnOffset, final Object[] values);

  /**
   * Sets the row having the specified index with the provided values. Sets the format of the edited
   * cells to the specified {@link SSCellFormat}.
   *
   * @param rowIndex of the row
   * @param values the values that has to be set
   * @param cellFormat
   * @see SSCellFormat
   */
  public void setRow(final int rowIndex, final Object[] values, final SSCellFormat cellFormat);

  /**
   * Set the row having the specified index with the provided values starting from the column
   * {@code columnOffset}. Sets the format of the edited cells to the specified
   * {@link SSCellFormat}.
   * 
   * @param rowIndex of the row
   * @param columnOffset
   * @param values the string values that has to be set
   * @param cellFormat
   * @see SSCellFormat
   */
  public void setRow(final int rowIndex, final int columnOffset, final Object[] values,
      final SSCellFormat cellFormat);
}
