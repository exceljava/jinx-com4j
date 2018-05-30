package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSortOrder implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlAscending(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDescending(2),
  ;

  private final int value;
  XlSortOrder(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
