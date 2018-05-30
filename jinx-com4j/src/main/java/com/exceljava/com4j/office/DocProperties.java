package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum DocProperties implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  offPropertyTypeNumber(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  offPropertyTypeBoolean(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  offPropertyTypeDate(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  offPropertyTypeString(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  offPropertyTypeFloat(5),
  ;

  private final int value;
  DocProperties(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
