package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlObjectSize implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlFitToPage(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlFullPage(3),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlScreenSize(1),
  ;

  private final int value;
  XlObjectSize(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
