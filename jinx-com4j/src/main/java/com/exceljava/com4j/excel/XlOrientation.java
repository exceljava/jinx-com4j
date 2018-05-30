package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4170
   * </p>
   */
  xlDownward(-4170),
  /**
   * <p>
   * The value of this constant is -4128
   * </p>
   */
  xlHorizontal(-4128),
  /**
   * <p>
   * The value of this constant is -4171
   * </p>
   */
  xlUpward(-4171),
  /**
   * <p>
   * The value of this constant is -4166
   * </p>
   */
  xlVertical(-4166),
  ;

  private final int value;
  XlOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
