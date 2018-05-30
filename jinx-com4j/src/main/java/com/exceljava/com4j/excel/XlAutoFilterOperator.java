package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlAutoFilterOperator implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlAnd(1),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlBottom10Items(4),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlBottom10Percent(6),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlOr(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlTop10Items(3),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlTop10Percent(5),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlFilterValues(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlFilterCellColor(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlFilterFontColor(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlFilterIcon(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlFilterDynamic(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlFilterNoFill(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlFilterAutomaticFontColor(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlFilterNoIcon(14),
  ;

  private final int value;
  XlAutoFilterOperator(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
