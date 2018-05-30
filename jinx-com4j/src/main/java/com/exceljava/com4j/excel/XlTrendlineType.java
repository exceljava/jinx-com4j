package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTrendlineType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlExponential(5),
  /**
   * <p>
   * The value of this constant is -4132
   * </p>
   */
  xlLinear(-4132),
  /**
   * <p>
   * The value of this constant is -4133
   * </p>
   */
  xlLogarithmic(-4133),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlMovingAvg(6),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPolynomial(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlPower(4),
  ;

  private final int value;
  XlTrendlineType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
