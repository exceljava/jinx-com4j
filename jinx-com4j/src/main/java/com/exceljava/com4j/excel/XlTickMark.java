package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTickMark implements ComEnum {
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlTickMarkCross(4),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTickMarkInside(2),
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlTickMarkNone(-4142),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlTickMarkOutside(3),
  ;

  private final int value;
  XlTickMark(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
