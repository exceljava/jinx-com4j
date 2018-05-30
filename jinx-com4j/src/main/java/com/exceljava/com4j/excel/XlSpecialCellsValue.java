package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSpecialCellsValue implements ComEnum {
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlErrors(16),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlLogical(4),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlNumbers(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTextValues(2),
  ;

  private final int value;
  XlSpecialCellsValue(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
