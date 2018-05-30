package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSaveAction implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDoNotSaveChanges(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSaveChanges(1),
  ;

  private final int value;
  XlSaveAction(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
