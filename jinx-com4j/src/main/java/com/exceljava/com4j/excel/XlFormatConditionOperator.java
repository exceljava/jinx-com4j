package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFormatConditionOperator implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlBetween(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlNotBetween(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlEqual(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlNotEqual(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlGreater(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlLess(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlGreaterEqual(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlLessEqual(8),
  ;

  private final int value;
  XlFormatConditionOperator(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
