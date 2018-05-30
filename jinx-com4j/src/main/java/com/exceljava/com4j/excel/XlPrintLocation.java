package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPrintLocation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPrintSheetEnd(1),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlPrintInPlace(16),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlPrintNoComments(-4142),
  ;

  private final int value;
  XlPrintLocation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
