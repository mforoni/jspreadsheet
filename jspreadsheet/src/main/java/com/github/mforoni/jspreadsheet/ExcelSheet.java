package com.github.mforoni.jspreadsheet;

import com.google.common.annotations.Beta;

/**
 * @author Foroni Marco
 */
@Beta
abstract class ExcelSheet implements Sheet {
  /**
   * Hide the column having the specified {@code columnIndex}.
   * 
   * @param columnIndex the index of the column to hide
   */
  public abstract void hideColumn(final int columnIndex);
}
