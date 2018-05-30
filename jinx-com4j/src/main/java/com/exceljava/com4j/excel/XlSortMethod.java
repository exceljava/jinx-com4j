package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSortMethod implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPinYin(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlStroke(2),
  ;

  private final int value;
  XlSortMethod(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
