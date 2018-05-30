package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSlicerCrossFilterType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSlicerNoCrossFilter(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSlicerCrossFilterShowItemsWithDataAtTop(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSlicerCrossFilterShowItemsWithNoData(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlSlicerCrossFilterHideButtonsWithNoData(4),
  ;

  private final int value;
  XlSlicerCrossFilterType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
