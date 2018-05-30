package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlEnableCancelKey implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlDisabled(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlErrorHandler(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlInterrupt(1),
  ;

  private final int value;
  XlEnableCancelKey(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
