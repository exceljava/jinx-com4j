package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlStdColorScale implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlColorScaleRYG(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlColorScaleGYR(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlColorScaleBlackWhite(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlColorScaleWhiteBlack(4),
  ;

  private final int value;
  XlStdColorScale(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
