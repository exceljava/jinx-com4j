package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDataSeriesType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlAutoFill(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlChronological(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlGrowth(2),
  /**
   * <p>
   * The value of this constant is -4132
   * </p>
   */
  xlDataSeriesLinear(-4132),
  ;

  private final int value;
  XlDataSeriesType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
