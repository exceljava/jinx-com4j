package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlLegendPosition implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4107
   * </p>
   */
  xlLegendPositionBottom(-4107),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlLegendPositionCorner(2),
  /**
   * <p>
   * The value of this constant is -4131
   * </p>
   */
  xlLegendPositionLeft(-4131),
  /**
   * <p>
   * The value of this constant is -4152
   * </p>
   */
  xlLegendPositionRight(-4152),
  /**
   * <p>
   * The value of this constant is -4160
   * </p>
   */
  xlLegendPositionTop(-4160),
  /**
   * <p>
   * The value of this constant is -4161
   * </p>
   */
  xlLegendPositionCustom(-4161),
  ;

  private final int value;
  XlLegendPosition(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
