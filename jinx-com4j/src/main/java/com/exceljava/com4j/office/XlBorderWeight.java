package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlBorderWeight implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlHairline(1),
  /**
   * <p>
   * The value of this constant is -4138
   * </p>
   */
  xlMedium(-4138),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlThick(4),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlThin(2),
  ;

  private final int value;
  XlBorderWeight(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
