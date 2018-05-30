package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTextOrientationMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTextOrientationHorizontal(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTextOrientationUpward(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoTextOrientationDownward(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoTextOrientationVerticalFarEast(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoTextOrientationVertical(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoTextOrientationHorizontalRotatedFarEast(6),
  ;

  private final int value;
  MsoTextOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
