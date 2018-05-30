package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSortOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSortRows(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSortColumns(1),
  ;

  private final int value;
  XlSortOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
