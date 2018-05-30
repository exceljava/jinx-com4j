package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlDisplayUnit implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  xlHundreds(-2),
  /**
   * <p>
   * The value of this constant is -3
   * </p>
   */
  xlThousands(-3),
  /**
   * <p>
   * The value of this constant is -4
   * </p>
   */
  xlTenThousands(-4),
  /**
   * <p>
   * The value of this constant is -5
   * </p>
   */
  xlHundredThousands(-5),
  /**
   * <p>
   * The value of this constant is -6
   * </p>
   */
  xlMillions(-6),
  /**
   * <p>
   * The value of this constant is -7
   * </p>
   */
  xlTenMillions(-7),
  /**
   * <p>
   * The value of this constant is -8
   * </p>
   */
  xlHundredMillions(-8),
  /**
   * <p>
   * The value of this constant is -9
   * </p>
   */
  xlThousandMillions(-9),
  /**
   * <p>
   * The value of this constant is -10
   * </p>
   */
  xlMillionMillions(-10),
  /**
   * <p>
   * The value of this constant is -4114
   * </p>
   */
  xlDisplayUnitCustom(-4114),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlDisplayUnitNone(-4142),
  ;

  private final int value;
  XlDisplayUnit(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
