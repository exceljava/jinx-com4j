package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSearchOrder implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlByColumns(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlByRows(1),
  ;

  private final int value;
  XlSearchOrder(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
