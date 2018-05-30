package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlEnableSelection implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlNoRestrictions(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlUnlockedCells(1),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlNoSelection(-4142),
  ;

  private final int value;
  XlEnableSelection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
