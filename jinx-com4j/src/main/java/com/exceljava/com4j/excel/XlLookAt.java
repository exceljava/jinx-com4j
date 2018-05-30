package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlLookAt implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPart(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlWhole(1),
  ;

  private final int value;
  XlLookAt(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
