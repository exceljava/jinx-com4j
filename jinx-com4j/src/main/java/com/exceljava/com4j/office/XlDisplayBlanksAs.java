package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlDisplayBlanksAs implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlInterpolated(3),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlNotPlotted(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlZero(2),
  ;

  private final int value;
  XlDisplayBlanksAs(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
