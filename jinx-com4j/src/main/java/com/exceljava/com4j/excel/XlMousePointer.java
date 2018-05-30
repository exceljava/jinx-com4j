package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlMousePointer implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlIBeam(3),
  /**
   * <p>
   * The value of this constant is -4143
   * </p>
   */
  xlDefault(-4143),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlNorthwestArrow(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlWait(2),
  ;

  private final int value;
  XlMousePointer(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
