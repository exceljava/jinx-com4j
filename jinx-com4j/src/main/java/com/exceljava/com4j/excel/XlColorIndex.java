package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlColorIndex implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlColorIndexAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlColorIndexNone(-4142),
  ;

  private final int value;
  XlColorIndex(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
