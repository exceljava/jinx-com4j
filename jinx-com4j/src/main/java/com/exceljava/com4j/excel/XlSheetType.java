package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSheetType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4109
   * </p>
   */
  xlChart(-4109),
  /**
   * <p>
   * The value of this constant is -4116
   * </p>
   */
  xlDialogSheet(-4116),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlExcel4IntlMacroSheet(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlExcel4MacroSheet(3),
  /**
   * <p>
   * The value of this constant is -4167
   * </p>
   */
  xlWorksheet(-4167),
  ;

  private final int value;
  XlSheetType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
