package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlApplyNamesOrder implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlColumnThenRow(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlRowThenColumn(1),
  ;

  private final int value;
  XlApplyNamesOrder(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
