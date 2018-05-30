package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoCalloutDropType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoCalloutDropMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoCalloutDropCustom(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoCalloutDropTop(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoCalloutDropCenter(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoCalloutDropBottom(4),
  ;

  private final int value;
  MsoCalloutDropType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
