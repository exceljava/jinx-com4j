package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSummaryReportType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4148
   * </p>
   */
  xlSummaryPivotTable(-4148),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlStandardSummary(1),
  ;

  private final int value;
  XlSummaryReportType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
