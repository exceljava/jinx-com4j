package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlWindowState implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4137
   * </p>
   */
  xlMaximized(-4137),
  /**
   * <p>
   * The value of this constant is -4140
   * </p>
   */
  xlMinimized(-4140),
  /**
   * <p>
   * The value of this constant is -4143
   * </p>
   */
  xlNormal(-4143),
  ;

  private final int value;
  XlWindowState(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
