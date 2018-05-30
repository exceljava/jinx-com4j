package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPresetLightingDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPresetLightingDirectionMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLightingTopLeft(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLightingTop(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLightingTopRight(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoLightingLeft(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoLightingNone(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoLightingRight(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoLightingBottomLeft(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoLightingBottom(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoLightingBottomRight(9),
  ;

  private final int value;
  MsoPresetLightingDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
