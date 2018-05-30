package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTickLabelPosition implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4127
   * </p>
   */
  xlTickLabelPositionHigh(-4127),
  /**
   * <p>
   * The value of this constant is -4134
   * </p>
   */
  xlTickLabelPositionLow(-4134),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlTickLabelPositionNextToAxis(4),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlTickLabelPositionNone(-4142),
  ;

  private final int value;
  XlTickLabelPosition(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
