package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSortType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSortLabels(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSortValues(1),
  ;

  private final int value;
  XlSortType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
