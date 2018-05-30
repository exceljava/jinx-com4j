package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlErrorBarType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4114
   * </p>
   */
  xlErrorBarTypeCustom(-4114),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlErrorBarTypeFixedValue(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlErrorBarTypePercent(2),
  /**
   * <p>
   * The value of this constant is -4155
   * </p>
   */
  xlErrorBarTypeStDev(-4155),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlErrorBarTypeStError(4),
  ;

  private final int value;
  XlErrorBarType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
