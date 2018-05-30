package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlLinkType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlLinkTypeExcelLinks(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlLinkTypeOLELinks(2),
  ;

  private final int value;
  XlLinkType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
