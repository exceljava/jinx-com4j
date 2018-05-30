package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlActionType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlActionTypeUrl(1),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlActionTypeRowset(16),
  /**
   * <p>
   * The value of this constant is 128
   * </p>
   */
  xlActionTypeReport(128),
  /**
   * <p>
   * The value of this constant is 256
   * </p>
   */
  xlActionTypeDrillthrough(256),
  ;

  private final int value;
  XlActionType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
