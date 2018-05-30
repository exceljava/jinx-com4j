package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSummaryColumn implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4131
   * </p>
   */
  xlSummaryOnLeft(-4131),
  /**
   * <p>
   * The value of this constant is -4152
   * </p>
   */
  xlSummaryOnRight(-4152),
  ;

  private final int value;
  XlSummaryColumn(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
