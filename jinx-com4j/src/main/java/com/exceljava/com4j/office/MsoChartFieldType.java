package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoChartFieldType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoChartFieldBubbleSize(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoChartFieldCategoryName(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoChartFieldPercentage(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoChartFieldSeriesName(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoChartFieldValue(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoChartFieldFormula(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoChartFieldRange(7),
  ;

  private final int value;
  MsoChartFieldType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
