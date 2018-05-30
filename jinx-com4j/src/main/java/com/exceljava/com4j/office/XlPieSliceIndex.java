package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlPieSliceIndex implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlOuterCounterClockwisePoint(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlOuterCenterPoint(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlOuterClockwisePoint(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlMidClockwiseRadiusPoint(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlCenterPoint(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlMidCounterClockwiseRadiusPoint(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlInnerClockwisePoint(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlInnerCenterPoint(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlInnerCounterClockwisePoint(9),
  ;

  private final int value;
  XlPieSliceIndex(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
