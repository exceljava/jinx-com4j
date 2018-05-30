package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFormatConditionType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCellValue(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlExpression(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlColorScale(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlDatabar(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlTop10(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlIconSets(6),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlUniqueValues(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlTextString(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlBlanksCondition(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlTimePeriod(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlAboveAverageCondition(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlNoBlanksCondition(13),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlErrorsCondition(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlNoErrorsCondition(17),
  ;

  private final int value;
  XlFormatConditionType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
