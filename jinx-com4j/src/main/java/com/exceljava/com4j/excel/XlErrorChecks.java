package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlErrorChecks implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlEvaluateToError(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTextDate(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlNumberAsText(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlInconsistentFormula(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlOmittedCells(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlUnlockedFormulaCells(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlEmptyCellReferences(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlListDataValidation(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlInconsistentListFormula(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlMisleadingFormat(10),
  ;

  private final int value;
  XlErrorChecks(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
