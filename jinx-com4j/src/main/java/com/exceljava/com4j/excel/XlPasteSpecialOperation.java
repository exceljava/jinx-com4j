package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPasteSpecialOperation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPasteSpecialOperationAdd(2),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlPasteSpecialOperationDivide(5),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlPasteSpecialOperationMultiply(4),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlPasteSpecialOperationNone(-4142),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPasteSpecialOperationSubtract(3),
  ;

  private final int value;
  XlPasteSpecialOperation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
