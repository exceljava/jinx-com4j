package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoCalloutType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoCalloutMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoCalloutOne(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoCalloutTwo(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoCalloutThree(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoCalloutFour(4),
  ;

  private final int value;
  MsoCalloutType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
