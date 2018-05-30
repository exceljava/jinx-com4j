package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlPieSliceLocation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlHorizontalCoordinate(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlVerticalCoordinate(2),
  ;

  private final int value;
  XlPieSliceLocation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
