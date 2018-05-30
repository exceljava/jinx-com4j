package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoGradientStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoGradientMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoGradientHorizontal(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoGradientVertical(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoGradientDiagonalUp(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoGradientDiagonalDown(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoGradientFromCorner(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoGradientFromTitle(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoGradientFromCenter(7),
  ;

  private final int value;
  MsoGradientStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
