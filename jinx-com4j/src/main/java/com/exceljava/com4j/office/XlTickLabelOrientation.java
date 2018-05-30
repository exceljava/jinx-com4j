package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlTickLabelOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlTickLabelOrientationAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is -4170
   * </p>
   */
  xlTickLabelOrientationDownward(-4170),
  /**
   * <p>
   * The value of this constant is -4128
   * </p>
   */
  xlTickLabelOrientationHorizontal(-4128),
  /**
   * <p>
   * The value of this constant is -4171
   * </p>
   */
  xlTickLabelOrientationUpward(-4171),
  /**
   * <p>
   * The value of this constant is -4166
   * </p>
   */
  xlTickLabelOrientationVertical(-4166),
  ;

  private final int value;
  XlTickLabelOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
