package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlRangeValueDataType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlRangeValueDefault(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlRangeValueXMLSpreadsheet(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlRangeValueMSPersistXML(12),
  ;

  private final int value;
  XlRangeValueDataType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
