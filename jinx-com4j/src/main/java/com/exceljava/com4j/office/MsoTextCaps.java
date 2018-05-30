package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextCaps implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoCapsMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoNoCaps(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSmallCaps(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAllCaps(2),
  ;

  private final int value;
  MsoTextCaps(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
