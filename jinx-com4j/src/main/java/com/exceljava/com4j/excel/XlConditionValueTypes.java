package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlConditionValueTypes implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlConditionValueNone(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlConditionValueNumber(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlConditionValueLowestValue(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlConditionValueHighestValue(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlConditionValuePercent(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlConditionValueFormula(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlConditionValuePercentile(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlConditionValueAutomaticMin(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlConditionValueAutomaticMax(7),
  ;

  private final int value;
  XlConditionValueTypes(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
