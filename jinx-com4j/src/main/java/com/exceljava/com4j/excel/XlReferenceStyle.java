package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlReferenceStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlA1(1),
  /**
   * <p>
   * The value of this constant is -4150
   * </p>
   */
  xlR1C1(-4150),
  ;

  private final int value;
  XlReferenceStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
