package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCreator implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1480803660
   * </p>
   */
  xlCreatorCode(1480803660),
  ;

  private final int value;
  XlCreator(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
