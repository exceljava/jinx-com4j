package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPlatform implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlMacintosh(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlMSDOS(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlWindows(2),
  ;

  private final int value;
  XlPlatform(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
