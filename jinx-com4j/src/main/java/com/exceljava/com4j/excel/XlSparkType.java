package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSparkType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSparkLine(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSparkColumn(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSparkColumnStacked100(3),
  ;

  private final int value;
  XlSparkType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
