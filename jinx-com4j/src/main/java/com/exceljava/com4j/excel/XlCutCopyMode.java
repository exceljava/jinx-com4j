package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCutCopyMode implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCopy(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCut(2),
  ;

  private final int value;
  XlCutCopyMode(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
