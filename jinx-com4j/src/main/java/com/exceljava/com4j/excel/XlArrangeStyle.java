package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlArrangeStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlArrangeStyleCascade(7),
  /**
   * <p>
   * The value of this constant is -4128
   * </p>
   */
  xlArrangeStyleHorizontal(-4128),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlArrangeStyleTiled(1),
  /**
   * <p>
   * The value of this constant is -4166
   * </p>
   */
  xlArrangeStyleVertical(-4166),
  ;

  private final int value;
  XlArrangeStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
