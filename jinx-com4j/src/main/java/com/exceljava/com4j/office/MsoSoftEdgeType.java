package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSoftEdgeType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoSoftEdgeTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoSoftEdgeTypeNone(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSoftEdgeType1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSoftEdgeType2(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoSoftEdgeType3(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoSoftEdgeType4(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoSoftEdgeType5(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoSoftEdgeType6(6),
  ;

  private final int value;
  MsoSoftEdgeType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
