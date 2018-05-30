package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlUnderlineStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4119
   * </p>
   */
  xlUnderlineStyleDouble(-4119),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlUnderlineStyleDoubleAccounting(5),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlUnderlineStyleNone(-4142),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlUnderlineStyleSingle(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlUnderlineStyleSingleAccounting(4),
  ;

  private final int value;
  XlUnderlineStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
