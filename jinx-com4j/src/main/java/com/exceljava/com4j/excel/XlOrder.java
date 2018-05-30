package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlOrder implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlDownThenOver(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlOverThenDown(2),
  ;

  private final int value;
  XlOrder(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
