package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPropertyDisplayedIn implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlDisplayPropertyInPivotTable(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDisplayPropertyInTooltip(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlDisplayPropertyInPivotTableAndTooltip(3),
  ;

  private final int value;
  XlPropertyDisplayedIn(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
