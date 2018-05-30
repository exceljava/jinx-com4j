package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDataLabelSeparator implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlDataLabelSeparatorDefault(1),
  ;

  private final int value;
  XlDataLabelSeparator(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
