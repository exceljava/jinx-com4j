package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlMarkerStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlMarkerStyleAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlMarkerStyleCircle(8),
  /**
   * <p>
   * The value of this constant is -4115
   * </p>
   */
  xlMarkerStyleDash(-4115),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlMarkerStyleDiamond(2),
  /**
   * <p>
   * The value of this constant is -4118
   * </p>
   */
  xlMarkerStyleDot(-4118),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlMarkerStyleNone(-4142),
  /**
   * <p>
   * The value of this constant is -4147
   * </p>
   */
  xlMarkerStylePicture(-4147),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlMarkerStylePlus(9),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlMarkerStyleSquare(1),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlMarkerStyleStar(5),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlMarkerStyleTriangle(3),
  /**
   * <p>
   * The value of this constant is -4168
   * </p>
   */
  xlMarkerStyleX(-4168),
  ;

  private final int value;
  XlMarkerStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
