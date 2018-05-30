package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlEditionFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlBIFF(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPICT(1),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlRTF(4),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlVALU(8),
  ;

  private final int value;
  XlEditionFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
