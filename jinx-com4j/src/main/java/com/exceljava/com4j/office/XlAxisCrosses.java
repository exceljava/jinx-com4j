package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlAxisCrosses implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlAxisCrossesAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is -4114
   * </p>
   */
  xlAxisCrossesCustom(-4114),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAxisCrossesMaximum(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlAxisCrossesMinimum(4),
  ;

  private final int value;
  XlAxisCrosses(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
