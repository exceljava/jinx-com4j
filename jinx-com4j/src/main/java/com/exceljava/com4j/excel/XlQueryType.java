package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlQueryType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlODBCQuery(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDAORecordset(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlWebQuery(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlOLEDBQuery(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlTextImport(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlADORecordset(7),
  ;

  private final int value;
  XlQueryType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
