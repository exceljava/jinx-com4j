package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFileAccess implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlReadOnly(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlReadWrite(2),
  ;

  private final int value;
  XlFileAccess(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
