package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPresetExtrusionDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPresetExtrusionDirectionMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoExtrusionBottomRight(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoExtrusionBottom(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoExtrusionBottomLeft(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoExtrusionRight(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoExtrusionNone(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoExtrusionLeft(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoExtrusionTopRight(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoExtrusionTop(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoExtrusionTopLeft(9),
  ;

  private final int value;
  MsoPresetExtrusionDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
