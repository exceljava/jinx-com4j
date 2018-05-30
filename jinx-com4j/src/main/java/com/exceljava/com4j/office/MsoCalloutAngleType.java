package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoCalloutAngleType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoCalloutAngleMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoCalloutAngleAutomatic(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoCalloutAngle30(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoCalloutAngle45(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoCalloutAngle60(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoCalloutAngle90(5),
  ;

  private final int value;
  MsoCalloutAngleType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
