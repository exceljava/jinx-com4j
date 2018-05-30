package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDeleteShiftDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4159
   * </p>
   */
  xlShiftToLeft(-4159),
  /**
   * <p>
   * The value of this constant is -4162
   * </p>
   */
  xlShiftUp(-4162),
  ;

  private final int value;
  XlDeleteShiftDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
