package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTextVisualLayoutType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlTextVisualLTR(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTextVisualRTL(2),
  ;

  private final int value;
  XlTextVisualLayoutType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
