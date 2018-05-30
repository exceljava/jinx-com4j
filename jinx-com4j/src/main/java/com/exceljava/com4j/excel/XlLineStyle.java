package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlLineStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlContinuous(1),
  /**
   * <p>
   * The value of this constant is -4115
   * </p>
   */
  xlDash(-4115),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlDashDot(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlDashDotDot(5),
  /**
   * <p>
   * The value of this constant is -4118
   * </p>
   */
  xlDot(-4118),
  /**
   * <p>
   * The value of this constant is -4119
   * </p>
   */
  xlDouble(-4119),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlSlantDashDot(13),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlLineStyleNone(-4142),
  ;

  private final int value;
  XlLineStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
