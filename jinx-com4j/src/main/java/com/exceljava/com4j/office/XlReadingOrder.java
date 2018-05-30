package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlReadingOrder implements ComEnum {
  /**
   * <p>
   * The value of this constant is -5002
   * </p>
   */
  xlContext(-5002),
  /**
   * <p>
   * The value of this constant is -5003
   * </p>
   */
  xlLTR(-5003),
  /**
   * <p>
   * The value of this constant is -5004
   * </p>
   */
  xlRTL(-5004),
  ;

  private final int value;
  XlReadingOrder(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
