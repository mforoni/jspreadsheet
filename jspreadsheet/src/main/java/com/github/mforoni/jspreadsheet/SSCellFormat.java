package com.github.mforoni.jspreadsheet;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.concurrent.Immutable;

/**
 * @author Foroni Marco
 */
@Immutable
public final class SSCellFormat {
  @Nonnull
  private final SSFont ssFont;
  @Nullable
  private final SSColor backgroundColor;

  public SSCellFormat(final SSFont ssFont) {
    this.ssFont = ssFont;
    this.backgroundColor = null;
  }

  public SSCellFormat(final SSFont ssFont, @Nullable final SSColor backgroundColor) {
    this.ssFont = ssFont;
    this.backgroundColor = backgroundColor;
  }

  @Nonnull
  public SSFont getSSFont() {
    return ssFont;
  }

  @Nullable
  public SSColor getBackgroundColour() {
    return backgroundColor;
  }
}
