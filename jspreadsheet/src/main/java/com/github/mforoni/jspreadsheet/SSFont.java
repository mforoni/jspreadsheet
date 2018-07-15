package com.github.mforoni.jspreadsheet;

import javax.annotation.concurrent.Immutable;
import jxl.write.WritableFont;
import jxl.write.WriteException;

/**
 * @author Foroni Marco
 */
@Immutable
public final class SSFont {
  public static final double DEFAULT_SIZE = 10;
  public static final String ARIAL = "Arial";
  public static final String CALIBRI = "Calibri";
  public static final SSFont ARIAL_10 = new SSFont(ARIAL, 10);
  public static final SSFont CALIBRI_10 = new SSFont(CALIBRI, 10);
  public static final SSFont BOLD_ARIAL_10 = new SSFont(ARIAL, 10, true);
  private final String name; // font name
  private final double size;
  private final boolean bold;

  private SSFont(final String name, final double size, final boolean bold) {
    this.name = name;
    this.size = size;
    this.bold = bold;
  }

  public SSFont(final String name) {
    this(name, DEFAULT_SIZE, false);
  }

  public SSFont(final String name, final double size) {
    this(name, size, false);
  }

  /**
   * Returns a new {@link WritableFont}.
   * 
   * @return
   * @throws WriteException
   */
  public WritableFont newWritableFont() throws WriteException {
    final WritableFont writableFont =
        new WritableFont(WritableFont.createFont(name), (int) Math.round(size));
    if (isBold()) {
      writableFont.setBoldStyle(WritableFont.BOLD);
    }
    return writableFont;
  }

  public static class Builder {
    private final String name;
    private final double size;
    private boolean bold;

    public Builder(final String name, final double pointSize) {
      this.name = name;
      this.size = pointSize;
      this.bold = false;
    }

    public Builder bold() {
      this.bold = true;
      return this;
    }

    public SSFont build() {
      return new SSFont(name, size, bold);
    }
  }

  public String getName() {
    return name;
  }

  public double getSize() {
    return size;
  }

  public boolean isBold() {
    return bold;
  }

  @Override
  public String toString() {
    return "SSFont [name=" + name + ", size=" + size + ", bold=" + bold + "]";
  }
}
