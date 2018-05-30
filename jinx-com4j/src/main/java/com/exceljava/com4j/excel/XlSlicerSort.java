package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSlicerSort implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSlicerSortDataSourceOrder(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSlicerSortAscending(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSlicerSortDescending(3),
  ;

  private final int value;
  XlSlicerSort(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
