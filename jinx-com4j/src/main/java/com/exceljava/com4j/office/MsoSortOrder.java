package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSortOrder implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSortOrderAscending(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSortOrderDescending(2),
  ;

  private final int value;
  MsoSortOrder(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
