package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlInsertShiftDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4121
   * </p>
   */
  xlShiftDown(-4121),
  /**
   * <p>
   * The value of this constant is -4161
   * </p>
   */
  xlShiftToRight(-4161),
  ;

  private final int value;
  XlInsertShiftDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
