package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPivotFieldRepeatLabels implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlDoNotRepeatLabels(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlRepeatLabels(2),
  ;

  private final int value;
  XlPivotFieldRepeatLabels(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
