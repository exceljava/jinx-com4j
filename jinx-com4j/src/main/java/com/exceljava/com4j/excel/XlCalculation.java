package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCalculation implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4105
   * </p>
   */
  xlCalculationAutomatic(-4105),
  /**
   * <p>
   * The value of this constant is -4135
   * </p>
   */
  xlCalculationManual(-4135),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCalculationSemiautomatic(2),
  ;

  private final int value;
  XlCalculation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
