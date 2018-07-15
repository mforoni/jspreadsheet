package com.github.mforoni.jspreadsheet;

import java.awt.Color;
import javax.annotation.concurrent.Immutable;
import org.apache.poi.ss.usermodel.IndexedColors;
import jxl.format.Colour;

/**
 * @author Foroni Marco
 */
@Immutable
public final class SSColor {
  // public static final SSColor AQUA = new SSColor(Colour.AQUA, IndexedColors.AQUA.getIndex(),
  // Color.AQUA);
  public static final SSColor BLACK =
      new SSColor(Colour.BLACK, IndexedColors.BLACK.getIndex(), Color.BLACK);
  public static final SSColor BLUE =
      new SSColor(Colour.BLUE, IndexedColors.BLUE.getIndex(), Color.BLUE);
  public static final SSColor GRAY_50 =
      new SSColor(Colour.GRAY_50, IndexedColors.GREY_50_PERCENT.getIndex(), Color.GRAY);
  public static final SSColor GRAY_25 =
      new SSColor(Colour.GRAY_25, IndexedColors.GREY_25_PERCENT.getIndex(), Color.LIGHT_GRAY);
  public static final SSColor GRAY_80 =
      new SSColor(Colour.GRAY_80, IndexedColors.GREY_80_PERCENT.getIndex(), Color.DARK_GRAY);
  public static final SSColor GREEN =
      new SSColor(Colour.GREEN, IndexedColors.GREEN.getIndex(), Color.GREEN);
  public static final SSColor ORANGE =
      new SSColor(Colour.ORANGE, IndexedColors.ORANGE.getIndex(), Color.ORANGE);
  public static final SSColor RED =
      new SSColor(Colour.RED, IndexedColors.RED.getIndex(), Color.RED);
  public static final SSColor YELLOW =
      new SSColor(Colour.YELLOW, IndexedColors.YELLOW.getIndex(), Color.YELLOW);
  // public static final SSColor LIGHT_BLUE = new SSColor(Colour.LIGHT_BLUE,
  // IndexedColors.LIGHT_BLUE.getIndex(),
  // Color.BLUE.brighter());
  // public static final SSColor DARK_BLUE = new SSColor(Colour.DARK_BLUE,
  // IndexedColors.DARK_BLUE.getIndex(),
  // Color.BLUE.darker());
  // public static final SSColor LIGHT_GREEN = new SSColor(Colour.LIGHT_GREEN,
  // IndexedColors.LIGHT_GREEN.getIndex(),
  // Color.GREEN.brighter());
  // public static final SSColor LIGHT_ORANGE = new SSColor(Colour.LIGHT_ORANGE,
  // IndexedColors.LIGHT_ORANGE.getIndex(), Color.ORANGE.brighter());
  // public static final SSColor LIGHT_TURQUOISE = new SSColor(Colour.LIGHT_TURQUOISE,
  // IndexedColors.LIGHT_TURQUOISE.getIndex(), null);
  private final Colour colour; // JXL
  private final short poiIndex; // POI
  private final java.awt.Color color; // ODS

  public SSColor(final Colour colour, final short poiIndex, final Color color) {
    this.colour = colour;
    this.poiIndex = poiIndex;
    this.color = color;
  }

  /**
   * Returns the JXL color.
   * 
   * @return
   * @see Colour
   */
  public Colour getColour() {
    return colour;
  }

  /**
   * Returns the POI color.
   * 
   * @return
   * @see IndexedColors
   */
  public short getPoiIndex() {
    return poiIndex;
  }

  /**
   * Returns the ODS color.
   * 
   * @return
   */
  public Color getColor() {
    return color;
  }

  @Override
  public String toString() {
    return "SSColor [colour=" + colour + ", poiIndex=" + poiIndex + "]";
  }
}
