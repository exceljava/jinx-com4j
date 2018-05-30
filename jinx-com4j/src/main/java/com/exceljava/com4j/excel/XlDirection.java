package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4121
   * </p>
   */
  xlDown(-4121),
  /**
   * <p>
   * The value of this constant is -4159
   * </p>
   */
  xlToLeft(-4159),
  /**
   * <p>
   * The value of this constant is -4161
   * </p>
   */
  xlToRight(-4161),
  /**
   * <p>
   * The value of this constant is -4162
   * </p>
   */
  xlUp(-4162),
  ;

  private final int value;
  XlDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
