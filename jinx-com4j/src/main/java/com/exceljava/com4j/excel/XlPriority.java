package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPriority implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4127
   * </p>
   */
  xlPriorityHigh(-4127),
  /**
   * <p>
   * The value of this constant is -4134
   * </p>
   */
  xlPriorityLow(-4134),
  /**
   * <p>
   * The value of this constant is -4143
   * </p>
   */
  xlPriorityNormal(-4143),
  ;

  private final int value;
  XlPriority(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
