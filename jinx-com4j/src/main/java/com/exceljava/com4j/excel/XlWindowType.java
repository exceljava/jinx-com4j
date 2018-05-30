package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlWindowType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlChartAsWindow(5),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlChartInPlace(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlClipboard(3),
  /**
   * <p>
   * The value of this constant is -4129
   * </p>
   */
  xlInfo(-4129),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlWorkbook(1),
  ;

  private final int value;
  XlWindowType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
