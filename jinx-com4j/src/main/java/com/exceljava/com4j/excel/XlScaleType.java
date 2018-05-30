package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlScaleType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4132
   * </p>
   */
  xlScaleLinear(-4132),
  /**
   * <p>
   * The value of this constant is -4133
   * </p>
   */
  xlScaleLogarithmic(-4133),
  ;

  private final int value;
  XlScaleType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
