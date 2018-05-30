package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDVAlertStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlValidAlertStop(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlValidAlertWarning(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlValidAlertInformation(3),
  ;

  private final int value;
  XlDVAlertStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
