package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPageBreakExtent implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPageBreakFull(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPageBreakPartial(2),
  ;

  private final int value;
  XlPageBreakExtent(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
