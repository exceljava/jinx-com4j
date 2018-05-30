package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlWBATemplate implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4109
   * </p>
   */
  xlWBATChart(-4109),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlWBATExcel4IntlMacroSheet(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlWBATExcel4MacroSheet(3),
  /**
   * <p>
   * The value of this constant is -4167
   * </p>
   */
  xlWBATWorksheet(-4167),
  ;

  private final int value;
  XlWBATemplate(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
