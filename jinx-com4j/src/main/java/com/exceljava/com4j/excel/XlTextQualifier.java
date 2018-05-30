package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTextQualifier implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlTextQualifierDoubleQuote(1),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlTextQualifierNone(-4142),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTextQualifierSingleQuote(2),
  ;

  private final int value;
  XlTextQualifier(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
