package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoOrgChartLayoutType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoOrgChartLayoutMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoOrgChartLayoutStandard(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoOrgChartLayoutBothHanging(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoOrgChartLayoutLeftHanging(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoOrgChartLayoutRightHanging(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoOrgChartLayoutDefault(5),
  ;

  private final int value;
  MsoOrgChartLayoutType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
