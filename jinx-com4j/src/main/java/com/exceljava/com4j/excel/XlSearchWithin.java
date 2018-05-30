package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSearchWithin implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlWithinSheet(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlWithinWorkbook(2),
  ;

  private final int value;
  XlSearchWithin(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
