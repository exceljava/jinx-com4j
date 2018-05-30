package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlBackground implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlBackgroundAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlBackgroundOpaque(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlBackgroundTransparent(2),
  ;

  private final int value;
  XlBackground(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
