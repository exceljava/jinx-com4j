package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSortMethodOld implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCodePage(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSyllabary(1),
  ;

  private final int value;
  XlSortMethodOld(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
