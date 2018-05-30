package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlArrowHeadWidth implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4138
   * </p>
   */
  xlArrowHeadWidthMedium(-4138),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlArrowHeadWidthNarrow(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlArrowHeadWidthWide(3),
  ;

  private final int value;
  XlArrowHeadWidth(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
