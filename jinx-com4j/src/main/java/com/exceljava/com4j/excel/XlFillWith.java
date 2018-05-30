package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFillWith implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4104
   * </p>
   */
  xlFillWithAll(-4104),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlFillWithContents(2),
  /**
   * <p>
   * The value of this constant is -4122
   * </p>
   */
  xlFillWithFormats(-4122),
  ;

  private final int value;
  XlFillWith(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
