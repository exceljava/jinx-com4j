package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlAxisGroup implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPrimary(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSecondary(2),
  ;

  private final int value;
  XlAxisGroup(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
