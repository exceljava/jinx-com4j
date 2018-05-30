package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlWindowView implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlNormalView(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPageBreakPreview(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPageLayoutView(3),
  ;

  private final int value;
  XlWindowView(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
