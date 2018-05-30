package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlErrorBarInclude implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlErrorBarIncludeBoth(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlErrorBarIncludeMinusValues(3),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlErrorBarIncludeNone(-4142),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlErrorBarIncludePlusValues(2),
  ;

  private final int value;
  XlErrorBarInclude(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
