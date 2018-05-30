package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTopBottom implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlTop10Top(1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlTop10Bottom(0),
  ;

  private final int value;
  XlTopBottom(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
