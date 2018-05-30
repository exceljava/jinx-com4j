package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlArrowHeadStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlArrowHeadStyleClosed(3),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlArrowHeadStyleDoubleClosed(5),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlArrowHeadStyleDoubleOpen(4),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlArrowHeadStyleNone(-4142),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlArrowHeadStyleOpen(2),
  ;

  private final int value;
  XlArrowHeadStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
