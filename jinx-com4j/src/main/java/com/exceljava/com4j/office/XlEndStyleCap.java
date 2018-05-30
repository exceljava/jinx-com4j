package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlEndStyleCap implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCap(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlNoCap(2),
  ;

  private final int value;
  XlEndStyleCap(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
