package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlRunAutoMacro implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlAutoActivate(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAutoClose(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlAutoDeactivate(4),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlAutoOpen(1),
  ;

  private final int value;
  XlRunAutoMacro(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
