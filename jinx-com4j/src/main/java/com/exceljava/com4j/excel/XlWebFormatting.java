package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlWebFormatting implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlWebFormattingAll(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlWebFormattingRTF(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlWebFormattingNone(3),
  ;

  private final int value;
  XlWebFormatting(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
