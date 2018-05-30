package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlBordersIndex implements ComEnum {
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlInsideHorizontal(12),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlInsideVertical(11),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlDiagonalDown(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlDiagonalUp(6),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlEdgeBottom(9),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlEdgeLeft(7),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlEdgeRight(10),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlEdgeTop(8),
  ;

  private final int value;
  XlBordersIndex(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
