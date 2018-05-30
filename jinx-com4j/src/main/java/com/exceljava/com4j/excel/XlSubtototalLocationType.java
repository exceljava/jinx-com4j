package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSubtototalLocationType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlAtTop(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAtBottom(2),
  ;

  private final int value;
  XlSubtototalLocationType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
