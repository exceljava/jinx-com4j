package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlWebSelectionType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlEntirePage(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAllTables(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSpecifiedTables(3),
  ;

  private final int value;
  XlWebSelectionType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
