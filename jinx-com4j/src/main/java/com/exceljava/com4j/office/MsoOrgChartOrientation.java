package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoOrgChartOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoOrgChartOrientationMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoOrgChartOrientationVertical(1),
  ;

  private final int value;
  MsoOrgChartOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
