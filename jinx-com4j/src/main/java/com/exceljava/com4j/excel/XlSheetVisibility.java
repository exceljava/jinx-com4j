package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSheetVisibility implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlSheetVisible(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlSheetHidden(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSheetVeryHidden(2),
  ;

  private final int value;
  XlSheetVisibility(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
