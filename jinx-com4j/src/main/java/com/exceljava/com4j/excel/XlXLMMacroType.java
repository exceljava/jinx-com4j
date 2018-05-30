package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlXLMMacroType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCommand(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlFunction(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlNotXLM(3),
  ;

  private final int value;
  XlXLMMacroType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
