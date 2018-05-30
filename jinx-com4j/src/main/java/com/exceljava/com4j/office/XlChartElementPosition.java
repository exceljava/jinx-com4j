package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlChartElementPosition implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlChartElementPositionAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is -4114
   * </p>
   */
  xlChartElementPositionCustom(-4114),
  ;

  private final int value;
  XlChartElementPosition(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
