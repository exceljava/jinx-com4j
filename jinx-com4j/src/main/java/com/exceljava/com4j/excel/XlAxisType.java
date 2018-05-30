package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlAxisType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCategory(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSeriesAxis(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlValue(2),
  ;

  private final int value;
  XlAxisType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
