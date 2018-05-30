package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSeriesNameLevel implements ComEnum {
  /**
   * <p>
   * The value of this constant is -3
   * </p>
   */
  xlSeriesNameLevelNone(-3),
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  xlSeriesNameLevelCustom(-2),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlSeriesNameLevelAll(-1),
  ;

  private final int value;
  XlSeriesNameLevel(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
