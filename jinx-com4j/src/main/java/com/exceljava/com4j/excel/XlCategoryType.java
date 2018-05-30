package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCategoryType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCategoryScale(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlTimeScale(3),
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlAutomaticScale(-4105),
  ;

  private final int value;
  XlCategoryType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
