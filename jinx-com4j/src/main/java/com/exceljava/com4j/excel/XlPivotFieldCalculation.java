package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPivotFieldCalculation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDifferenceFrom(2),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlIndex(9),
  /**
   * <p>
   * The value of this constant is -4143
   * </p>
   */
  xlNoAdditionalCalculation(-4143),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlPercentDifferenceFrom(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPercentOf(3),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlPercentOfColumn(7),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlPercentOfRow(6),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlPercentOfTotal(8),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlRunningTotal(5),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlPercentOfParentRow(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlPercentOfParentColumn(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlPercentOfParent(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlPercentRunningTotal(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlRankAscending(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlRankDecending(15),
  ;

  private final int value;
  XlPivotFieldCalculation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
