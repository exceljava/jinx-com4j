package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFindLookIn implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4123
   * </p>
   */
  xlFormulas(-4123),
  /**
   * <p>
   * The value of this constant is -4144
   * </p>
   */
  xlComments(-4144),
  /**
   * <p>
   * The value of this constant is -4163
   * </p>
   */
  xlValues(-4163),
  ;

  private final int value;
  XlFindLookIn(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
