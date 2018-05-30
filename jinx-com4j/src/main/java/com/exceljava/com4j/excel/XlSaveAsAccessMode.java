package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSaveAsAccessMode implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlExclusive(3),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlNoChange(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlShared(2),
  ;

  private final int value;
  XlSaveAsAccessMode(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
