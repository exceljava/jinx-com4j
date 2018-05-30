package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPageOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlLandscape(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPortrait(1),
  ;

  private final int value;
  XlPageOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
