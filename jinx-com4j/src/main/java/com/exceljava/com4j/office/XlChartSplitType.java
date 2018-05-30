package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlChartSplitType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSplitByPosition(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSplitByPercentValue(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlSplitByCustomSplit(4),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSplitByValue(2),
  ;

  private final int value;
  XlChartSplitType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
