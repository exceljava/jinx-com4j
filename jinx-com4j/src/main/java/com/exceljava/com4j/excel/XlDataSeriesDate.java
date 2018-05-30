package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDataSeriesDate implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlDay(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlMonth(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlWeekday(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlYear(4),
  ;

  private final int value;
  XlDataSeriesDate(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
