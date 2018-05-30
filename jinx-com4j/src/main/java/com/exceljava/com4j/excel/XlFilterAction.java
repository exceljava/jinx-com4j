package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFilterAction implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlFilterCopy(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlFilterInPlace(1),
  ;

  private final int value;
  XlFilterAction(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
