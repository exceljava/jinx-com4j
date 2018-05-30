package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoDocProperties implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoPropertyTypeNumber(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoPropertyTypeBoolean(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoPropertyTypeDate(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoPropertyTypeString(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoPropertyTypeFloat(5),
  ;

  private final int value;
  MsoDocProperties(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
