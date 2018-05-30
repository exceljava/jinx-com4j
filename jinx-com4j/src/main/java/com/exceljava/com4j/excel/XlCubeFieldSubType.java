package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCubeFieldSubType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCubeHierarchy(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCubeMeasure(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlCubeSet(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlCubeAttribute(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlCubeCalculatedMeasure(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlCubeKPIValue(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlCubeKPIGoal(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlCubeKPIStatus(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlCubeKPITrend(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlCubeKPIWeight(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlCubeImplicitMeasure(11),
  ;

  private final int value;
  XlCubeFieldSubType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
