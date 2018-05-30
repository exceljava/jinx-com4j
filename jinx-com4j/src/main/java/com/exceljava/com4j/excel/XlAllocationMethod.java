package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlAllocationMethod implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlEqualAllocation(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlWeightedAllocation(2),
  ;

  private final int value;
  XlAllocationMethod(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
