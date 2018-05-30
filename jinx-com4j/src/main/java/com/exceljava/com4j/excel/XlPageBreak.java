package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPageBreak implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlPageBreakAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is -4135
   * </p>
   */
  xlPageBreakManual(-4135),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlPageBreakNone(-4142),
  ;

  private final int value;
  XlPageBreak(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
