package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPivotFieldDataType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDate(2),
  /**
   * <p>
   * The value of this constant is -4145
   * </p>
   */
  xlNumber(-4145),
  /**
   * <p>
   * The value of this constant is -4158
   * </p>
   */
  xlText(-4158),
  ;

  private final int value;
  XlPivotFieldDataType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
