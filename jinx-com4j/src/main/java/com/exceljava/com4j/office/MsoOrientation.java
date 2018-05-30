package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoOrientationMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoOrientationHorizontal(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoOrientationVertical(2),
  ;

  private final int value;
  MsoOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
