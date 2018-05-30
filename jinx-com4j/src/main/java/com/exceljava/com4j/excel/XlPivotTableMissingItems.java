package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPivotTableMissingItems implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlMissingItemsDefault(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlMissingItemsNone(0),
  /**
   * <p>
   * The value of this constant is 32500
   * </p>
   */
  xlMissingItemsMax(32500),
  /**
   * <p>
   * The value of this constant is 1048576
   * </p>
   */
  xlMissingItemsMax2(1048576),
  ;

  private final int value;
  XlPivotTableMissingItems(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
