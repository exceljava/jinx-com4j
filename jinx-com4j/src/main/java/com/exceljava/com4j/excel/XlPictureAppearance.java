package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPictureAppearance implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPrinter(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlScreen(1),
  ;

  private final int value;
  XlPictureAppearance(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
