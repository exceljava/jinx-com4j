package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlReferenceType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlAbsolute(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAbsRowRelColumn(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlRelative(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlRelRowAbsColumn(3),
  ;

  private final int value;
  XlReferenceType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
