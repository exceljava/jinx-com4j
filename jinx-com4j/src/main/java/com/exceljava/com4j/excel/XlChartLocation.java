package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlChartLocation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlLocationAsNewSheet(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlLocationAsObject(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlLocationAutomatic(3),
  ;

  private final int value;
  XlChartLocation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
