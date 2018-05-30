package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCommandUnderlines implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlCommandUnderlinesAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is -4146
   * </p>
   */
  xlCommandUnderlinesOff(-4146),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCommandUnderlinesOn(1),
  ;

  private final int value;
  XlCommandUnderlines(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
