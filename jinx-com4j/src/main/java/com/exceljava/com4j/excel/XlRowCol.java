package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlRowCol implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlColumns(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlRows(1),
  ;

  private final int value;
  XlRowCol(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
