package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSearchDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlNext(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPrevious(2),
  ;

  private final int value;
  XlSearchDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
