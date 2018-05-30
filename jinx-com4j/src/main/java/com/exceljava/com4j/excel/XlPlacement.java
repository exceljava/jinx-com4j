package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPlacement implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlFreeFloating(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlMove(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlMoveAndSize(1),
  ;

  private final int value;
  XlPlacement(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
