package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLineDashStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoLineDashStyleMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLineSolid(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLineSquareDot(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLineRoundDot(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoLineDash(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoLineDashDot(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoLineDashDotDot(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoLineLongDash(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoLineLongDashDot(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoLineLongDashDotDot(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoLineSysDash(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoLineSysDot(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoLineSysDashDot(12),
  ;

  private final int value;
  MsoLineDashStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
