package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPivotTableVersionList implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlPivotTableVersion2000(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPivotTableVersion10(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPivotTableVersion11(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPivotTableVersion12(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlPivotTableVersion14(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlPivotTableVersion15(5),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlPivotTableVersionCurrent(-1),
  ;

  private final int value;
  XlPivotTableVersionList(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
