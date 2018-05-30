package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlErrorBarDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4168
   * </p>
   */
  xlX(-4168),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlY(1),
  ;

  private final int value;
  XlErrorBarDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
