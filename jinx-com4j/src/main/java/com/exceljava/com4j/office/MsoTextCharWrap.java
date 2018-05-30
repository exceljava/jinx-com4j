package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextCharWrap implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoCharWrapMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoNoCharWrap(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoStandardCharWrap(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoStrictCharWrap(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoCustomCharWrap(3),
  ;

  private final int value;
  MsoTextCharWrap(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
