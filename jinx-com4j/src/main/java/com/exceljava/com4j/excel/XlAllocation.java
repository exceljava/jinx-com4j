package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlAllocation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlManualAllocation(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAutomaticAllocation(2),
  ;

  private final int value;
  XlAllocation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
