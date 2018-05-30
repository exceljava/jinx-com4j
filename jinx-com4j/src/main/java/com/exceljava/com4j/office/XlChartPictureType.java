package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlChartPictureType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlStackScale(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlStack(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlStretch(1),
  ;

  private final int value;
  XlChartPictureType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
