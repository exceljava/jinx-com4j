package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlArrowHeadLength implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlArrowHeadLengthLong(3),
  /**
   * <p>
   * The value of this constant is -4138
   * </p>
   */
  xlArrowHeadLengthMedium(-4138),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlArrowHeadLengthShort(1),
  ;

  private final int value;
  XlArrowHeadLength(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
