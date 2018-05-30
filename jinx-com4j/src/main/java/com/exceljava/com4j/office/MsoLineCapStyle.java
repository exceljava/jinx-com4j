package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLineCapStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoLineCapMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLineCapSquare(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLineCapRound(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLineCapFlat(3),
  ;

  private final int value;
  MsoLineCapStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
