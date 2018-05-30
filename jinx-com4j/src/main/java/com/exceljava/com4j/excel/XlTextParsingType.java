package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTextParsingType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlDelimited(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlFixedWidth(2),
  ;

  private final int value;
  XlTextParsingType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
