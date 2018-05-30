package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlAllocationValue implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlAllocateValue(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAllocateIncrement(2),
  ;

  private final int value;
  XlAllocationValue(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
