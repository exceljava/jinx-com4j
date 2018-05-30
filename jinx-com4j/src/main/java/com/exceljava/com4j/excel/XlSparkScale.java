package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSparkScale implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSparkScaleGroup(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSparkScaleSingle(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSparkScaleCustom(3),
  ;

  private final int value;
  XlSparkScale(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
